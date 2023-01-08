from OpenServer import OpenServer
import pandas as pd
import xlwings as xw
import scipy.optimize as optim
import numpy as np


def calculate_vlp(OSC, method_name='brent'):

    df = get_data(r'data\Day3.VirtualFlow.xlsx', 'VirtualFlowmeter')
    OSC.do_command('GAP.MOD[{PROD}].WELL[{W1}].UNMASK()')
    lb = 500.
    ub = 10000.
    for idx, row in df.iterrows():
        if row['Active Rows'] == 'Enabled':

            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].TPD.SensVarCalcValue[1]', row['Water Cut'])
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].TPD.SensVarCalcValue[2]', row['Gas Oil Ratio'])
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].TPD.SensVarCalcValue[3]', row['Tubing Head Pressure'])
            OSC.do_command('GAP.MOD[{PROD}].WELL[{W1}].TPD.Calc()')
            try:
                if method_name == 'brent':
                    root = optim.brentq(lambda x: calc_vlp(x, OSC, row['Gauge Pressure']), lb, ub, rtol=1.e-3,
                                        full_output=True)
                elif method_name == 'newton':
                    root = optim.newton(lambda x: calc_vlp(x, OSC, row['Gauge Pressure']), 0.5 * (lb + ub), rtol=1.e-3,
                                        full_output=True)
                else:
                    root = optim.bisect(lambda x: calc_vlp(x, OSC, row['Gauge Pressure']), lb, ub, rtol=1.e-3,
                                        full_output=True)

            except ValueError:
                root = optim.newton(lambda x: calc_vlp(x, OSC, row['Gauge Pressure']), ub, rtol=1.e-3,
                                full_output=True)

            df.at[idx, 'VLP'] = root[0]
            df.at[idx, 'Number iteration'] = root[1].function_calls
            df.at[idx, 'PetroleumExperts5 Ex'] = float(
                OSC.get_value('GAP.MOD[{PROD}].WELL[{W1}].TPD.CalcVarResult[0]'))


    name_pairs = [('P6', 'VLP'), ('T6', 'PetroleumExperts5 Ex'), ('Y6', 'Number iteration')]
    export_to_excel(df, r'data\Day3.VirtualFlow.xlsx', 'VirtualFlowmeter', name_pairs)
    OSC.do_command('GAP.MOD[{PROD}].WELL[{W1}].MASK()')


def calc_vlp(x, OSC, press_match):

    OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].TPD.SensVarCalcValue[0]', x)

    OSC.do_command('GAP.MOD[{PROD}].WELL[{W1}].TPD.Calc()')
    press_calc = float(OSC.get_value('GAP.MOD[{PROD}].WELL[{W1}].TPD.CalcVarResult[2]'))

    return press_calc - press_match


def calculate_ipr(OSC):

    df = get_data(r'data\Day3.VirtualFlow.xlsx', 'VirtualFlowmeter')
    OSC.do_command('GAP.MOD[{PROD}].WELL[{W1}].UNMASK()')

    for idx, row in df.iterrows():
        if row['Active Rows'] == 'Enabled' and not(pd.isnull(row['PetroleumExperts5 Ex'])):
            pres = (row['PetroleumExperts5 Ex'] - 1.) / 0.0689476
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].IPR[0].ResPres', row['Reservoir Pressure'])
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].IPR[0].PI', row['PI'])
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].IPR[0].WCT', row['Water Cut'])
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].IPR[0].GOR', row['Gas Oil Ratio'])
            OSC.do_command(f'GAP.MOD[{{PROD}}].WELL[{{W1}}].IPR[0].IPRCalc({pres},0)')
            df.at[idx, 'IPR'] = float(OSC.get_value('GAP.MOD[{PROD}].WELL[{W1}].IPR[0].IPRCalcResult'))

    export_to_excel(df, r'data\Day3.VirtualFlow.xlsx', 'VirtualFlowmeter', [('Q6', 'IPR')])
    OSC.do_command('GAP.MOD[{PROD}].WELL[{W1}].MASK()')


def calculate_vlpipr(OSC):

    df = get_data(r'data\Day3.VirtualFlow.xlsx', 'VirtualFlowmeter')
    OSC.do_command('GAP.MOD[{PROD}].WELL[{W1}].UNMASK()')
    for idx, row in df.iterrows():
        if row['Active Rows'] == 'Enabled':
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].MeasuredResP', row['Reservoir Pressure'])
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].IPR[0].PI', row['PI'])
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].MeasuredWCT', row['Water Cut'])
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].MeasuredGOR', row['Gas Oil Ratio'])
            OSC.set_value('GAP.MOD[{PROD}].WELL[{W1}].MeasuredPres', row['Tubing Head Pressure'])
            OSC.do_command('GAP.WELLCALC(MOD[0].WELL[{W1}])')
            df.at[idx, 'VLP/IPR'] = float(OSC.get_value('GAP.MOD[{PROD}].WELL[{W1}].EstimatedLiqRate'))

    export_to_excel(df, r'data\Day3.VirtualFlow.xlsx', 'VirtualFlowmeter', [('R6', 'VLP/IPR')])
    OSC.do_command('GAP.MOD[{PROD}].WELL[{W1}].MASK()')


def calculate_choke(OSC):

    df = get_data(r'data\Day3.VirtualFlow.xlsx', 'VirtualFlowmeter')
    OSC.do_command('GAP.MOD[{PROD}].WELL[{W1}].MASK()')
    for idx, row in df.iterrows():
        if row['Active Rows'] == 'Enabled':
            OSC.set_value('GAP.MOD[{PROD}].SEP[{Sep1}].SolverPres[0]', row['Flow Line Pressure'])

            OSC.set_value('GAP.MOD[{PROD}].FluidList[{TestFluid}].WCT', row['Water Cut'])
            OSC.set_value('GAP.MOD[{PROD}].FluidList[{TestFluid}].GOR', row['Gas Oil Ratio'])

            OSC.set_value('GAP.MOD[{PROD}].SOURCE[{Source1}].PRESSURE', row['Tubing Head Pressure'])
            OSC.set_value('GAP.MOD[{PROD}].SOURCE[{Source1}].Temperature', row['Tubing Head Temperature'])

            OSC.set_value('GAP.MOD[{PROD}].INLCHK[{InlChk1}].ChokeDiameter', row['Choke Size'])

            OSC.do_command('GAP.SOLVENETWORK(0, MOD[0])')
            df.at[idx, 'Choke'] = float(OSC.get_value('GAP.MOD[{PROD}].SOURCE[{Source1}].SolverResults[0].LiqRate'))

    export_to_excel(df, r'data\Day3.VirtualFlow.xlsx', 'VirtualFlowmeter', [('S6', 'Choke')])


def get_data(name, sh_name):

    return pd.read_excel(name, sheet_name=sh_name, header=0)[4:].reset_index()


def export_to_excel(data, wb_name, sh_name, name_pairs):

    sh = xw.Book(wb_name).sheets[sh_name]
    for pair in name_pairs:
        sh.range(pair[0]).options(index=False, header=False).value = data.loc[:, pair[1]]
    xw.Book(wb_name).save()


if __name__ == '__main__':

    with OpenServer() as OSC:
        calculate_vlp(OSC, 'bisect')
        calculate_vlp(OSC, 'newton')
        calculate_vlp(OSC, 'brent')
        calculate_ipr(OSC)
        calculate_vlpipr(OSC)
        calculate_choke(OSC)


