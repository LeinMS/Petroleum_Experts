
from OpenServer import OpenServer
import pandas as pd
import xlwings as xw
import traceback


def calculate_gap(OSC):

    df_data = get_data(r'data\GAP_matching.xlsx', 'MatchGAP')

    get_from_gap(OSC, df_data, range(df_data.shape[0]))

    export_to_excel(df_data, r'data\GAP_matching.xlsx', 'MatchGAP', 'A5')


def get_from_gap(OSC, df_data, id_range):
    """
    Расчет модели GAP в цикле по выбранным срезам
    :param OSC: объект для коннекта с OpenServer
    :param df_data: dataframe с ихсодными и расчетными данными
    :param id_range: range или list с индексами, для которых выоплняется расчет
    :return:
    """
    for idx in id_range:
        row = df_data.iloc[idx]
        if row['Test']['Active Rows'] == 'Enabled':

            set_well_data(OSC, row['W1'], 'W1')
            set_well_data(OSC, row['W2'], 'W2')
            OSC.set_value('GAP.MOD[{PROD}].SEP[{Sep1}].SolverPres[0]', row['Separator']['Pressure'])

            OSC.do_command('GAP.SOLVENETWORK(0)')

            get_well_data(OSC, df_data, idx, 'W1')
            get_well_data(OSC, df_data, idx, 'W2')

            df_data.at[idx, ('Separator', 'Temperature Calc')] = float(OSC.get_value('GAP.MOD[{PROD}].SEP[{Sep1}].SolverResults[0].Temp'))
            mass_diff = float(OSC.get_value('GAP.MOD[{PROD}].SolverStatusList[0].MaxMassBalanceDiff'))
            if mass_diff < 0.2:
                df_data.at[idx, ('Separator', 'Calc status')] = 1
            else:
                df_data.at[idx, ('Separator', 'Calc status')] = 0


def set_well_data(OSC, row, w_name):

    OSC.set_value(f'GAP.MOD[{{PROD}}].WELL[{{{w_name}}}].AlqValue', row['Operating Frequency'])
    OSC.set_value(f'GAP.MOD[{{PROD}}].WELL[{{{w_name}}}].IPR[0].WCT', row['Water Cut'])
    OSC.set_value(f'GAP.MOD[{{PROD}}].WELL[{{{w_name}}}].IPR[0].ResPres', row['Reservoir Pressure'])


def get_well_data(OSC, df, idx, w_name):

    df.at[idx, (w_name, 'Liquid Rate Calc')] = float(OSC.get_value(f'GAP.MOD[{{PROD}}].WELL[{{{w_name}}}].SolverResults[0].LiqRate'))
    df.at[idx, (w_name, 'Tubing Head Pressure Calc')] = float(OSC.get_value(f'GAP.MOD[{{PROD}}].WELL[{{{w_name}}}].SolverResults[0].FWHP'))
    df.at[idx, (w_name, 'PIP calc')] = float(OSC.get_value(f'GAP.MOD[{{PROD}}].WELL[{{{w_name}}}].SolverResults[0].OtherResult[5]'))


def get_data(name, sh_name):

    df = pd.read_excel(name, sheet_name=sh_name, header=[0, 1])[2:].reset_index(drop=True)
    df.columns = df.columns.set_levels(df.columns.levels[0].astype(str), level=0)

    return df


def export_to_excel(data, wb_name, sh_name, start):

    sh = xw.Book(wb_name).sheets[sh_name]
    sh.range(start).options(index=False, header=False).value = data
    xw.Book(wb_name).save()


if __name__ == '__main__':

    with OpenServer() as OSC:
        calculate_gap(OSC)
