import pandas as pd


def read_bess_input(file_name, sheet_name):
    """Return BESS specifications as dictionary."""
    df = pd.read_excel(file_name, sheet_name=sheet_name, index_col=0)
    bess_dict = {key: [value, unit] for key, value, unit in zip(df.index, df['Values'], df['Units'])}
    return bess_dict


def read_price_input(file_name, half_hour_sheet, daily_sheet):
    """Return prices from all 3 markets as pandas Dataframe"""
    df_half_hour = pd.read_excel(file_name, sheet_name=half_hour_sheet, index_col=0)
    df_daily = pd.read_excel(file_name, sheet_name=daily_sheet, index_col=0)

    df_price = pd.concat([df_half_hour['Market 1 Price [£/MWh]'], df_half_hour['Market 2 Price [£/MWh]']], axis=1,
                         keys=[1, 2])
    df_price[3] = list(df_daily['Market 3 Price [£/MWh]'].repeat(48))

    df_price.reset_index(inplace=True)
    df_price.index += 1

    df = df_price.copy()
    df.drop(columns=['index'], inplace=True)
    price_dict = {(idx, n): df.at[idx, n] for idx in df.index for n in df.columns}
    return price_dict, df_price
