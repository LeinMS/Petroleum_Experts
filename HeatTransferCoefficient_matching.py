from OpenServer import OpenServer
import pandas as pd
import xlwings as xw
from data import export_to_excel, get_data
from datetime import datetime, timedelta
import traceback
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates


def fill_prosper(OSC):
    """
    занесение данных в секцию VLP/IPR
    :param OSC: объект   для коннекта с OpenServer
    :return:
    """
    df = get_data(r'data\results.xlsx', 'data')[1:]
    add_data_to_prosper(OSC, df)


def get_vlp_pres(OSC):

    corr_name = OSC.get_value('PROSPER.ANL.SYS.TubingLabel')
    df = get_dhgp_from_prosper(OSC, corr_name)
    export_to_excel(df, r'data\results.xlsx', 'data', 'O1')


def calculate_htc(OSC):
    """
    расчет htc в секции BHP from WHP и выгрузка в соответствующий лист excel файла
    :param OSC: объект   для коннекта с OpenServer
    :return:
    """

    df_tests = get_data(r'data\results.xlsx', 'data')[1:]
    df_tests['WC'] = df_tests.apply(lambda x: get_WC(x['LiquidRate'], x['WaterRate']), axis=1)
    df_slice = df_tests[['Date', 'WC', 'DHGT', 'FLT']]

    df = get_htc_from_prosper(OSC)

    df_slice = df_slice.merge(df, left_on='Date', right_on='Time').drop(columns='Time')
    export_to_excel(df_slice, r'data\results.xlsx', 'u_value', 'A1')


def calculate_corr(OSC):
    """
    расчет забойного давления для корреляций, указанных в шапке таблицы на листе 'corr'
    выгрузка в соответствующий лист excel файла
    :param OSC: объект   для коннекта с OpenServer
    :return:
    """
    df_tests = get_data(r'data\results.xlsx', 'data')[1:]
    df_slice = df_tests[['Date', 'DHGP']]

    corr_name = xw.Book(r'data\results.xlsx').sheets['corr'].range('A1').options(expand='right').value[2:]
    for c_n in corr_name:
        df_slice = df_slice.merge(get_dhgp_from_prosper(OSC, c_n), left_on='Date', right_on='Time').drop(columns='Time')

    export_to_excel(df_slice, r'data\results.xlsx', 'corr', 'A1')


def calculate_ipr(OSC):
    """
    расчет PI в секции IPR при помощи метода Vogel, используя забойное по лучшей корреляции (PE3)
    выгрузка в соответствующий лист excel файла
    :param OSC: объект   для коннекта с OpenServer
    :return:
    """
    corr_name = 'PetroleumExperts3'
    df_tests = get_data(r'data\results.xlsx', 'data')[1:]
    df_tests['WC'] = df_tests.apply(lambda x: get_WC(x['LiquidRate'], x['WaterRate']), axis=1)
    df_tests['GOR'] = df_tests.apply(lambda x: get_GOR(x['LiquidRate'], x['WaterRate'], x['GasRate']), axis=1)

    df_slice = df_tests[['Date', 'PR fill', 'LiquidRate', 'WC', 'GOR']]
    df_slice = df_slice.merge(get_dhgp_from_prosper(OSC, corr_name), left_on='Date', right_on='Time').drop(columns='Time')

    df = get_ipr_from_prosper(OSC, df_slice, corr_name)

    df_slice = df_slice.merge(df, left_on='Date', right_on='Time').drop(columns='Time')
    export_to_excel(df_slice, r'data\results.xlsx', 'ipr', 'A1')


def calculate_system(OSC):
    """
    расчет секции System и выгрузка Дебита жидкости, давления и температуры на манометре
    выгрузка в соответствующий лист excel файла
    :param OSC: объект   для коннекта с OpenServer
    :return:
    """
    df_tests = get_data(r'data\results.xlsx', 'data')[1:]
    df_tests['WC'] = df_tests.apply(lambda x: get_WC(x['LiquidRate'], x['WaterRate']), axis=1)
    df_tests['GOR'] = df_tests.apply(lambda x: get_GOR(x['LiquidRate'], x['WaterRate'], x['GasRate']), axis=1)

    df_slice = df_tests[['Date', 'PR fill', 'WC', 'GOR', 'FWHP', 'LiquidRate', 'DHGP', 'FLT']]
    df_slice = df_slice.merge(get_data(r'data\results.xlsx', 'ipr')[1:][['Date', 'PI']], left_on='Date', right_on='Date')

    df = get_system_from_prosper(OSC, df_slice)

    df_slice = df_slice.merge(df, left_on='Date', right_on='Time').drop(columns='Time')
    export_to_excel(df_slice, r'data\results.xlsx', 'system', 'A1')


get_WC = lambda liq, wat: 100. * wat / liq
get_GOR = lambda liq, wat, gas: 1000. * gas / (liq - wat)
date_to_days = lambda x: (x - datetime(1970, 1, 1)).days
days_to_date = lambda x: datetime(1970, 1, 1) + timedelta(days=x)


def get_system_from_prosper(OSC, df_data):
    """
    расчет секции System и выгрузка Дебита жидкости, давления и температуры на манометре
    :param OSC: объект   для коннекта с OpenServer
    :param df_data: dataframe с исходными данными
    :return: dataframe с результатми расчета
    """
    OSC.set_value('PROSPER.ANL.SYS.TubingLabel', 'PetroleumExperts3')
    OSC.set_value('PROSPER.Sin.IPR.Single.IprMethod', 0)
    df = pd.DataFrame(columns=['Time', 'LiquidRateCalc', 'DHGPCalc', 'THTcalc'])
    for idx, row in df_data.iterrows():
        is_correct_row = pd.isna(row['PR fill']) | pd.isna(row['WC']) | pd.isna(row['GOR']) | pd.isna(row['FWHP'])
        if not is_correct_row:
            OSC.set_value('PROSPER.Sin.IPR.Single.Pres', row['PR fill'])
            OSC.set_value('PROSPER.Sin.IPR.Single.Wc', row['WC'])
            OSC.set_value('PROSPER.Sin.IPR.Single.totgor', row['GOR'])
            OSC.set_value('PROSPER.Sin.IPR.Single.Pindex', row['PI'])

            OSC.set_value('PROSPER.ANL.SYS.Pres', row['FWHP'])
            OSC.set_value('PROSPER.ANL.SYS.WC', row['WC'])
            OSC.set_value('PROSPER.ANL.SYS.GOR', row['GOR'])

            OSC.do_command('PROSPER.ANL.SYS.Calc')

            new_row = pd.Series({'Time': row['Date'],
                                 'LiquidRateCalc': float(OSC.get_value('PROSPER.OUT.SYS.Results [0].Sol.LiqRate')),
                                 'DHGPCalc': float(OSC.get_value('PROSPER.OUT.SYS.Results[0].Sol.GaugeP[0]')),
                                 'THTcalc': float(OSC.get_value('PROSPER.OUT.SYS.Results [0].Sol.WHTemperature'))
                                })
            df = pd.concat([df, new_row.to_frame().T], ignore_index=True)

    return df


def get_ipr_from_prosper(OSC, df_data, corr_name):
    """
        расчет PI в секции IPR при помощи метода Vogel
        :param OSC: объект   для коннекта с OpenServer
        :param df_data: dataframe с исходными данными
        :param corr_name: имя корреляции
        :return: dataframe с результатми расчета
    """
    OSC.set_value('PROSPER.Sin.IPR.Single.IprMethod', 1)
    df = pd.DataFrame(columns=['Time', 'PI'])
    for idx, row in df_data.iterrows():
        is_correct_row = pd.isna(row['PR fill']) | pd.isna(row['WC']) | pd.isna(row['GOR']) | pd.isna(row['PetroleumExperts3']) | pd.isna(row['LiquidRate'])
        if not is_correct_row:
            OSC.set_value('PROSPER.Sin.IPR.Single.Pres', row['PR fill'])
            OSC.set_value('PROSPER.Sin.IPR.Single.Wc', row['WC'])
            OSC.set_value('PROSPER.Sin.IPR.Single.totgor', row['GOR'])
            OSC.set_value('PROSPER.Sin.IPR.Single.Ptest', row[corr_name])
            OSC.set_value('PROSPER.Sin.IPR.Single.Qtest', row['LiquidRate'])

            OSC.do_command('PROSPER.IPR.Calc')

            new_row = pd.Series({'Time': row['Date'],
                                 'PI': float(OSC.get_value('PROSPER.Sin.IPR.Single.PINSAV'))
                                })
            df = pd.concat([df, new_row.to_frame().T], ignore_index=True)

    return df


def get_dhgp_from_prosper(OSC, corr_name):
    """
     расчет забойного давления для корреляции в секции BHP from WHP
    :param OSC: объект   для коннекта с OpenServer
    :param corr_name: имя корреляции
    :return: dataframe с расчетом
    """
    OSC.set_value('PROSPER.ANL.WHP.TubingLabel', corr_name)
    OSC.do_command('PROSPER.ANL.WHP.CALC')

    df = pd.DataFrame(columns=['Time', corr_name])

    df['Time'] = [days_to_date(int(x)) for x in OSC.get_value(f'PROSPER.ANL.WHP.Data[$].Time').split('|')[:-1]]
    df[corr_name] = [float(x) for x in OSC.get_value(f'PROSPER.ANL.WHP.Data[$].BHP').split('|')[:-1]]

    count_htc = int(OSC.get_value('PROSPER.ANL.WHP.Data.Count'))
    for i in range(count_htc):
        days_from_prosper = int(float(OSC.get_value(f'PROSPER.ANL.WHP.Data[{i}].Time')))
        new_row = pd.Series({'Time': datetime(1970, 1, 1) + timedelta(days=days_from_prosper),
                             corr_name: float(OSC.get_value(f'PROSPER.ANL.WHP.Data[{i}].BHP'))
                            })
        df = pd.concat([df, new_row.to_frame().T], ignore_index=True)

    return df


def get_htc_from_prosper(OSC):
    """
    Используя секцию BHP from WHP, матчится HTC
    :param OSC: объект для коннекта с OpenServer
    :return: dataframe с расчитанным htc
    """
    OSC.do_command('PROSPER.ANL.WHP.CALC')

    df = pd.DataFrame(columns=['Time', 'HTC', 'WHTC_CALC'])

    df['Time'] = [days_to_date(int(x)) for x in OSC.get_value(f'PROSPER.ANL.WHP.Data[$].Time').split('|')[:-1]]
    df['HTC'] = [float(x) for x in OSC.get_value(f'PROSPER.ANL.WHP.Data[$].HTC').split('|')[:-1]]
    df['WHTC_CALC'] = [float(x) for x in OSC.get_value(f'PROSPER.ANL.WHP.Data[$].WHT').split('|')[:-1]]

    count_htc = int(OSC.get_value('PROSPER.ANL.WHP.Data.Count'))
    for i in range(count_htc):
        days_from_prosper = int(float(OSC.get_value(f'PROSPER.ANL.WHP.Data[{i}].Time')))
        new_row = pd.Series({'Time': datetime(1970, 1, 1) + timedelta(days=days_from_prosper),
                             'HTC': float(OSC.get_value(f'PROSPER.ANL.WHP.Data[{i}].HTC')),
                             'WHTC': float(OSC.get_value(f'PROSPER.ANL.WHP.Data[{i}].WHTC'))
                            })
        df = pd.concat([df, new_row.to_frame().T], ignore_index=True)

    return df


def add_data_to_prosper(OSC, df):

    clear_prosper_section(OSC, 'VMT')
    clear_prosper_section(OSC, 'WHP')
    for idx, row in df.iterrows():
        add_row_to_vlpipr(OSC, row)
        add_row_to_bhpwhp(OSC, row)


def clear_prosper_section(OSC, section_name):
    if int(OSC.get_value(f'PROSPER.ANL.{section_name}.Data.Count')) != 0:
        OSC.set_value(f'PROSPER.ANL.{section_name}.Data.RESET', '')


def add_row_to_vlpipr(OSC, row):

    nrows = int(OSC.get_value('PROSPER.ANL.VMT.Data.Count'))
    OSC.set_value(f'PROSPER.ANL.VMT.Data[{nrows}].Date', row['Date'])
    OSC.set_value(f'PROSPER.ANL.VMT.Data[{nrows}].THpres', row['FWHP'])
    OSC.set_value(f'PROSPER.ANL.VMT.Data[{nrows}].THtemp', row['FLT'])
    OSC.set_value(f'PROSPER.ANL.VMT.Data[{nrows}].WC', get_WC(row['LiquidRate'], row['WaterRate']))
    OSC.set_value(f'PROSPER.ANL.VMT.Data[{nrows}].Rate', row['LiquidRate'])
    OSC.set_value(f'PROSPER.ANL.VMT.Data[{nrows}].Gdepth', row['GD'])
    OSC.set_value(f'PROSPER.ANL.VMT.Data[{nrows}].Gpres', row['DHGP'])
    OSC.set_value(f'PROSPER.ANL.VMT.Data[{nrows}].Pres', row['PR fill'])
    OSC.set_value(f'PROSPER.ANL.VMT.Data[{nrows}].GOR', get_GOR(row['LiquidRate'], row['WaterRate'], row['GasRate']))
    OSC.set_value(f'PROSPER.ANL.VMT.Data[{nrows}].GORfree', 0.)


def add_row_to_bhpwhp(OSC, row):

    nrows = int(OSC.get_value('PROSPER.ANL.WHP.Data.Count'))

    OSC.set_value(f'PROSPER.ANL.WHP.Data[{nrows}].Time', date_to_days(row['Date']))
    OSC.set_value(f'PROSPER.ANL.WHP.Data[{nrows}].Rate', row['LiquidRate'])
    OSC.set_value(f'PROSPER.ANL.WHP.Data[{nrows}].WHP', row['FWHP'])
    OSC.set_value(f'PROSPER.ANL.WHP.Data[{nrows}].WHT', row['FLT'])
    OSC.set_value(f'PROSPER.ANL.WHP.Data[{nrows}].GASF', get_GOR(row['LiquidRate'], row['WaterRate'], row['GasRate']))
    OSC.set_value(f'PROSPER.ANL.WHP.Data[{nrows}].WATF', get_WC(row['LiquidRate'], row['WaterRate']))


def plot_wh_temperature(OSC):
    df_calc = get_htc_from_prosper(OSC)
    df = get_data(r'data\results.xlsx', 'data')[1:]

    fig = plt.figure()
    ax = fig.add_subplot()

    locator = mdates.AutoDateLocator(minticks=3, maxticks=7)
    formatter = mdates.ConciseDateFormatter(locator)
    ax.xaxis.set_major_formatter(formatter)
    ax.xaxis.set_major_locator(locator)

    p1 = ax.plot(df.loc[:, 'Date'], df.loc[:, 'FLT'], label='whtc')
    p1 = ax.plot(df_calc.loc[:, 'Time'], df_calc.loc[:, 'WHTC_CALC'], label='whtc_calc')
    fig.autofmt_xdate()

    ax.legend()

    plt.show()


if __name__ == '__main__':

    with OpenServer() as OSC:
        fill_prosper(OSC)
        plot_wh_temperature(OSC)
        calculate_htc(OSC)
        calculate_corr(OSC)
        calculate_ipr(OSC)
        calculate_system(OSC)
