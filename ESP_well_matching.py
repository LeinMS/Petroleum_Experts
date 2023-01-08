
from OpenServer import OpenServer
import pandas as pd
import xlwings as xw
import scipy.optimize as optim
from scipy.interpolate import interp1d
import numpy as np


def match_vlp(OSC, method_name='brent'):

    df_data = get_data(r'data\Day4_ESPWell#1.xlsx', 'esp_match')

    for idx, row in df_data.iterrows():
        if row['Active Rows'] == 'Enabled':
            df_out = get_vlp_from_prosper(OSC, row, method_name)
            export_to_excel(df_out, r'data\Day4_ESPWell#1.xlsx', 'esp_match', f'L{5 + idx}')
    xw.Book(r'data\Day4_ESPWell#1.xlsx').save()


def calculate_vlp(OSC):

    df_data = get_data(r'data\Day4_ESPWell#1.xlsx', 'esp_match')
    for idx, row in df_data.iterrows():
        if row['Active Rows'] == 'Enabled':
            df_out = calc_vlp_prosper(OSC, row)
            export_to_excel(df_out, r'data\Day4_ESPWell#1.xlsx', 'esp_match', f'M{5 + idx}')
    xw.Book(r'data\Day4_ESPWell#1.xlsx').save()


def match_ipr(OSC):

    df_data = get_data(r'data\Day4_ESPWell#1.xlsx', 'esp_match')
    OSC.set_value('PROSPER.SIN.IPR.Single.IprMethod', 1)
    for idx, row in df_data.iterrows():
        if row['Active Rows'] == 'Enabled':
            df_out = get_ipr_from_prosper(OSC, row)
            export_to_excel(df_out, r'data\Day4_ESPWell#1.xlsx', 'esp_match', f'O{5 + idx}')
    xw.Book(r'data\Day4_ESPWell#1.xlsx').save()


def calc_system(OSC):

    df_data = get_data(r'data\Day4_ESPWell#1.xlsx', 'esp_match')
    for idx, row in df_data.iterrows():
        if row['Active Rows'] == 'Enabled':
            df_out = get_system_from_prosper(OSC, row)
            export_to_excel(df_out, r'data\Day4_ESPWell#1.xlsx', 'esp_match', f'P{5 + idx}')
    xw.Book(r'data\Day4_ESPWell#1.xlsx').save()


def calc_vlp_prosper(OSC, row):

    set_vlp_to_gap(OSC, row)

    OSC.set_value('PROSPER.Sin.ESP.Wear', row['Pump Wear Factor Calc'])

    OSC.do_command('PROSPER.REFRESH')
    OSC.do_command('PROSPER.ANL.TCC.Calc')

    value_regime = OSC.get_value('PROSPER.OUT.TCC.Results[10].Regime[$]')
    index_pip = value_regime.count('|', 0, value_regime.index('-11')) + 1
    index_bhp = value_regime.count('|', 0, value_regime.index('100'))

    new_row = pd.Series({'PIP_Calc': float(OSC.get_value(f'PROSPER.OUT.TCC.Results[10].Pres[{index_pip}]')),
                         'BHP_Calc': float(OSC.get_value(f'PROSPER.OUT.TCC.Results[10].Pres[{index_bhp}]'))
                         }, dtype=float)

    return new_row.to_frame().T


def get_vlp_from_prosper(OSC, row, method_name):

    lb = -1.
    ub = 1.
    rtol = 0.001

    set_vlp_to_gap(OSC, row)

    OSC.do_command('PROSPER.REFRESH')
    OSC.do_command('PROSPER.ANL.TCC.Calc')

    value_regime = OSC.get_value('PROSPER.OUT.TCC.Results[10].Regime[$]')
    index_pip = value_regime.count('|', 0, value_regime.index('-11')) + 1
    index_bhp = value_regime.count('|', 0, value_regime.index('100'))

    if method_name == 'brent':
        root = optim.brentq(lambda x: calc_vlp(x, OSC, index_pip, row['PIP']), lb, ub, xtol=rtol, full_output=True)
    elif method_name == 'bisect':
        root = optim.bisect(lambda x: calc_vlp(x, OSC, index_pip, row['PIP']), lb, ub, xtol=rtol, full_output=True)
    else:
        root = interpolate_solve(lambda x: calc_vlp(x, OSC, index_pip, row['PIP']), lb, ub, tol=rtol)

    new_row = pd.Series({'PWF_Calc': root[0],
                         'PIP_Calc': float(OSC.get_value(f'PROSPER.OUT.TCC.Results[10].Pres[{index_pip}]')),
                         'BHP_Calc': float(OSC.get_value(f'PROSPER.OUT.TCC.Results[10].Pres[{index_bhp}]')),
                         'PI': np.NaN, 'Liq': np.NaN, 'PIP': np.NaN,
                         f'Number func calls': root[1].function_calls
                         }, dtype=float)


    return new_row.to_frame().T


def set_vlp_to_gap(OSC, row):

    OSC.set_value('PROSPER.ANL.TCC.Pres', row['Tubing Head Pressure'])
    OSC.set_value('PROSPER.ANL.TCC.WC', row['Water Cut'])
    OSC.set_value('PROSPER.ANL.TCC.Rate', row['Liquid Rate'])
    OSC.set_value('PROSPER.Sin.ESP.Frequency', row['Operating Frequency'])


def get_ipr_from_prosper(OSC, row):


    OSC.set_value('PROSPER.Sin.IPR.Single.Pres', row['Reservoir Pressure'])
    OSC.set_value('PROSPER.SIN.IPR.Single.Wc', row['Water Cut'])
    OSC.set_value('PROSPER.Sin.IPR.Single.Qtest', row['Liquid Rate'])
    OSC.set_value('PROSPER.Sin.IPR.Single.Ptest', row['Bottom Hole Pressure Calc'])

    OSC.do_command('PROSPER.REFRESH')
    OSC.do_command('PROSPER.IPR.Calc')

    new_row = pd.Series({'PI_Calc': float(OSC.get_value('PROSPER.Sin.IPR.Single.PINSAV'))
                        })

    return new_row.to_frame().T


def get_system_from_prosper(OSC, row):


    OSC.set_value('PROSPER.SIN.IPR.Single.IprMethod', 0)
    OSC.set_value('PROSPER.Sin.IPR.Single.Pres', row['Reservoir Pressure'])
    OSC.set_value('PROSPER.Sin.IPR.Single.Pindex', row['Productivity Index (PI) Calc'])
    OSC.set_value('PROSPER.ANL.SYS.Pres', row['Tubing Head Pressure'])
    OSC.set_value('PROSPER.ANL.SYS.WC', row['Water Cut'])
    OSC.set_value('PROSPER.Sin.ESP.Wear', row['Pump Wear Factor Calc'])
    OSC.set_value('PROSPER.Sin.ESP.Frequency', row['Operating Frequency'])

    OSC.do_command('PROSPER.ANL.SYS.Calc')

    new_row = pd.Series({'Liq_Calc': float(OSC.get_value('PROSPER.OUT.SYS.Results[0].Sol.LiqRate')),
                         'PIP_Calc': float(OSC.get_value('PROSPER.OUT.SYS.Results[0].Sol.PIP'))
                         })

    return new_row.to_frame().T


def calc_vlp(x, OSC, index_row, press_match):

    OSC.set_value('PROSPER.Sin.ESP.Wear', x)
    OSC.do_command('PROSPER.ANL.TCC.Calc')

    press_calc = float(OSC.get_value(f'PROSPER.OUT.TCC.Results[10].Pres[{index_row}]'))
    if int(OSC.get_value(f'PROSPER.OUT.TCC.Results[10].Regime[{index_row-1}]')) != -11:
        press_calc = 0.

    delta_p = press_calc - press_match
    print(delta_p)
    return delta_p


def get_data(name, sh_name):

    df = pd.read_excel(name, sheet_name=sh_name, header=1, usecols='A:Q')[2:].reset_index()

    return df


def export_to_excel(data, wb_name, sh_name, start):

    sh = xw.Book(wb_name).sheets[sh_name]
    sh.range(start).options(index=False, header=False).value = data


def interpolate_solve(f, lb, ub, max_iter=100, tol=1.e-3):

    fb = f(lb)
    fu = f(ub)
    x_new = ub
    iter = 2
    while iter < max_iter and abs((lb - ub)) > 2. * tol:

        x_new = interpolate_var([fb, fu], [lb, ub])
        f_new = f(x_new)

        if fb * f_new > 0.:
            lb = x_new
            fb = f_new
        else:
            ub = x_new
            fu = f_new

        iter += 1

    root = optim.RootResults(root=x_new, iterations=iter, function_calls=iter, flag=True)

    return [x_new, root]


def interpolate_var(x, y):

    return interp1d(x, y)(0)



if __name__ == '__main__':

    with OpenServer() as OSC:
        #match_vlp(OSC, 'bisect')
        match_vlp(OSC, 'brent')
        #match_vlp(OSC, 'interp')
        #calculate_vlp(OSC)
        match_ipr(OSC)
        calc_system(OSC)

