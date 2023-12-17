import time
import pyomo.environ as pyo
from pyomo.opt import SolverFactory
from datetime import datetime, timedelta
from postprocessing import *

M = 100_000     # arbitrarily large constant used to linearize the model


def main():
    bess_spec = read_bess_input('Second Round Technical Question - Attachment 1.xlsx', 'Data')
    price, df_price = read_price_input('Second Round Technical Question - Attachment 2.xlsx',
                                       'Half-hourly data', 'Daily data')
    timestep = len(df_price)
    market_total = len(df_price.columns) - 1

    initial_time = time.time()
    print(datetime.now().strftime("%I:%M:%S %p"), ': Preparing optimization model...')

    # model
    model = pyo.ConcreteModel()
    model.t = pyo.RangeSet(timestep)    # model time range is defined by range in given market prices
    model.n = pyo.RangeSet(market_total)   # index for 3 different markets

    # BESS parameters
    model.bess_capacity = pyo.Param(initialize=bess_spec['Max storage volume'][0])  # 4MWh
    model.bess_max_discharge_rate = pyo.Param(initialize=bess_spec['Max discharging rate'][0])  # 2MW
    model.bess_max_charge_rate = pyo.Param(initialize=bess_spec['Max charging rate'][0])    # 2MW
    model.bess_discharge_eff = pyo.Param(initialize=1-bess_spec['Battery discharging efficiency'][0])   # 0.05
    model.bess_charge_eff = pyo.Param(initialize=1-bess_spec['Battery charging efficiency'][0])   # 0.05
    model.bess_cycle_life = pyo.Param(initialize=bess_spec['Lifetime (2)'][0])    # 5000

    # market parameters
    model.market_prices = pyo.Param(model.t, model.n, initialize=price)

    # charge/discharge binary variables
    model.bi_discharge = pyo.Var(model.t, within=pyo.Binary)
    model.bi_charge = pyo.Var(model.t, within=pyo.Binary)

    # BESS optimisation variables
    model.bess_discharge_rate = pyo.Var(model.t, within=pyo.NonNegativeReals, bounds=(0, model.bess_max_discharge_rate))
    model.bess_charge_rate = pyo.Var(model.t, within=pyo.NonNegativeReals, bounds=(0, model.bess_max_charge_rate))
    model.soc = pyo.Var(model.t, within=pyo.NonNegativeReals, bounds=(0, model.bess_capacity))

    # market optimisation variables
    model.export_rate = pyo.Var(model.t, model.n, within=pyo.NonNegativeReals)
    model.import_rate = pyo.Var(model.t, model.n, within=pyo.NonNegativeReals)

    # constraints
    model.soc_change = pyo.Constraint(model.t, rule=soc_change_rule)
    model.export_sum = pyo.Constraint(model.t, rule=export_sum_rule)
    model.import_sum = pyo.Constraint(model.t, rule=import_sum_rule)
    model.discharge_binary = pyo.Constraint(model.t, rule=discharge_binary_rule)
    model.charge_binary = pyo.Constraint(model.t, rule=charge_binary_rule)
    model.mutual_exclusive = pyo.Constraint(model.t, rule=mutual_exclusive_constraint)  # only charge or discharge
    model.day_ahead_export = pyo.Constraint(model.t, rule=day_ahead_export_constraint)
    model.day_ahead_import = pyo.Constraint(model.t, rule=day_ahead_import_constraint)

    # objective function
    model.obj = pyo.Objective(rule=revenue, sense=pyo.maximize)

    model.pprint()
    runtime = time.time() - initial_time
    print('Time elapsed for model creation: ', str(timedelta(seconds=runtime)))

    # solver
    print(datetime.now().strftime("%I:%M:%S %p"), ": Solving for optimal solution...")
    initial_time = time.time()
    opt = SolverFactory('cbc')  # cbc or glpk
    opt.options['ratio'] = 0.02
    opt.options['sec'] = 1200
    opt.solve(model, tee=True, keepfiles=True, logfile='optimiser.log')
    runtime = time.time() - initial_time
    print('Solve time = ', str(timedelta(seconds=runtime)))

    df_raw = save_results_to_df(model)
    save_to_csv(df_raw, df_price)


def soc_change_rule(model, t):
    if t == 1:
        return model.soc[t] == model.bess_capacity - (model.bess_discharge_rate[t] * 0.5 / model.bess_discharge_eff) \
            + (model.bess_charge_rate[t] * 0.5 * model.bess_discharge_eff)
    else:
        return model.soc[t] == model.soc[t-1] - (model.bess_discharge_rate[t] * 0.5 / model.bess_discharge_eff) \
            + (model.bess_charge_rate[t] * 0.5 * model.bess_discharge_eff)


def export_sum_rule(model, t):
    return sum([model.export_rate[t, n] for n in model.n]) == model.bess_discharge_rate[t]


def import_sum_rule(model, t):
    return sum([model.import_rate[t, n] for n in model.n]) == model.bess_charge_rate[t]


def discharge_binary_rule(model, t):
    return model.bess_discharge_rate[t] <= model.bi_discharge[t] * M


def charge_binary_rule(model, t):
    return model.bess_charge_rate[t] <= model.bi_charge[t] * M


def mutual_exclusive_constraint(model, t):
    return model.bi_discharge[t] + model.bi_charge[t] == 1


def day_ahead_export_constraint(model, t):
    if (t-1) % 48 != 0:
        return model.export_rate[t, 3] == model.export_rate[((t - 1) // 48 * 48 + 1), 3]
    else:
        return pyo.Constraint.Skip


def day_ahead_import_constraint(model, t):
    if (t-1) % 48 != 0:
        return model.import_rate[t, 3] == model.import_rate[((t - 1) // 48 * 48 + 1), 3]
    else:
        return pyo.Constraint.Skip


def revenue(model):
    return 0.5 * (pyo.summation(model.market_prices, model.export_rate)
                  - pyo.summation(model.market_prices, model.import_rate))


def save_results_to_df(model):
    model_vars = model.component_map(ctype=pyo.Var)

    tmp_series = []  # to hold extracted values
    for k in model_vars.keys():
        v = model_vars[k]
        s = pd.Series(v.extract_values(), index=v.extract_values().keys())

        # unstack the series if series is multi-indexed
        if type(s.index[0]) == tuple:  # it is multi-indexed
            s = s.unstack(level=1)
        else:
            s = pd.DataFrame(s)

        s.columns = pd.MultiIndex.from_tuples([(k, t) for t in s.columns])
        tmp_series.append(s)

    df = pd.concat(tmp_series, axis=1)
    df.columns = df.columns.get_level_values(0)
    df.columns = ['discharge_indicator', 'charge_indicator', 'bess_discharge', 'bess_charge', 'soc', 'm1_export',
                  'm2_export', 'm3_export', 'm1_import', 'm2_import', 'm3_import']
    return df


if __name__ == '__main__':
    main()
