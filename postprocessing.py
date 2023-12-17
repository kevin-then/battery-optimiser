from preprocessing import *


def save_to_csv(model_result, price_param):
    bess_spec = read_bess_input('Second Round Technical Question - Attachment 1.xlsx', 'Data')

    df = pd.concat([model_result, price_param], axis=1)
    df.rename(columns={1: 'm1_price', 2: 'm2_price', 3: 'm3_price'}, inplace=True)
    df['soc'] = df['soc'] / bess_spec['Max storage volume'][0] * 100

    df['revenue'] = 0.5 * ((df['m1_export'] - df['m1_import']) * df['m1_price']
                           + (df['m2_export'] - df['m2_import']) * df['m2_price']
                           + (df['m3_export'] - df['m3_import']) * df['m3_price'])

    for col in ['total_revenue', 'capex', 'opex', 'simple_payback_yr', 'cycles_used', 'profit_year1', 'profit_year2',
                'profit_year3']:
        df.insert(len(df.columns), col, '')

    df.loc[1, 'capex'] = bess_spec['Capex'][0]
    df.loc[1, 'opex'] = bess_spec['Fixed Operational Costs'][0]
    df.loc[1, 'total_revenue'] = df['revenue'].sum()
    df.loc[1, 'simple_payback_yr'] = df.loc[1, 'capex'] / (df.loc[1, 'total_revenue'] - (df.loc[1, 'opex'] * 3)) * 3
    df.loc[1, 'cycles_used'] = df['bess_discharge'].sum() * 0.5 / bess_spec['Max storage volume'][0]
    df.loc[1, 'profit_year1'] = df[(df['index'] >= '2018-01-01') & (df['index'] < '2019-01-01')]['revenue'].sum()
    df.loc[1, 'profit_year2'] = df[(df['index'] >= '2019-01-01') & (df['index'] < '2020-01-01')]['revenue'].sum()
    df.loc[1, 'profit_year3'] = df[(df['index'] >= '2020-01-01') & (df['index'] < '2021-01-01')]['revenue'].sum()

    df.set_index('index', inplace=True)
    df.drop(columns=['discharge_indicator', 'charge_indicator'], inplace=True)
    df.to_csv('opt_output.csv')
