import traceback

from OpenServer import OpenServer
import pandas as pd
import xlwings as xw
import nlopt as nl
from GAP_calculate import get_from_gap, get_data, export_to_excel
import traceback


def match_temperature_gap_each_time(OSC):
    """
    адаптация модели GAP на температуру на сепараторе отдельно для каждого среза
    :param OSC: объект для коннекта с OpenServer
    :return:
    """
    f_path = r'data\GAP_matching.xlsx'
    sh_name = 'MatchGAP'
    df_data = get_data(f_path, sh_name)
    for idx, row in df_data.iterrows():
        match_temperature(OSC, df_data, [idx])

    export_to_excel(df_data, f_path, sh_name, 'A5')
    export_htc_to_excel(df_data, f_path, sh_name)


def match_temperature_gap(OSC):
    """
    адаптация модели GAP на температуру на сепараторе с постоянным значением параметров для всех срезов
    :param OSC: объект для коннекта с OpenServer
    :return:
    """
    f_path = r'data\GAP_matching.xlsx'
    sh_name = 'MatchGAP'
    df_data = get_data(f_path, sh_name)

    match_temperature(OSC, df_data, range(df_data.shape[0]))

    export_to_excel(df_data, f_path, sh_name, 'A5')
    export_htc_to_excel(df_data, f_path, sh_name)


def match_thp_gap_each_time(OSC):
    """
    адаптация модели GAP на буферное давление на скважинах отдельно для каждого среза
    :param OSC: объект для коннекта с OpenServer
    :return:
    """
    f_path = r'data\GAP_matching.xlsx'
    sh_name = 'MatchGAP'
    df_data = get_data(f_path, sh_name)
    for idx, row in df_data.iterrows():
        match_thp(OSC, df_data, [idx])

    export_to_excel(df_data, f_path, sh_name, 'A5')
    export_fc_to_excel(df_data, f_path, sh_name)


def match_thp_gap(OSC):
    """
     адаптация модели GAP на буферное давление на скважинах с постоянным значением параметров для всех срезов
     :param OSC: объект для коннекта с OpenServer
     :return:
     """
    f_path = r'data\GAP_matching.xlsx'
    sh_name = 'MatchGAP'
    df_data = get_data(f_path, sh_name)

    match_thp(OSC, df_data, range(df_data.shape[0]))

    export_to_excel(df_data, f_path, sh_name, 'A5')
    export_fc_to_excel(df_data, f_path, sh_name)


def match_temperature(OSC, df_data, id_range):

    name_prop = 'HTCSUR'
    labels = [['1', '3'], ['{Manifold}']]
    init = [45.431864 for i in range(len(labels))]
    lb = [10. for i in range(len(labels))]
    ub = [60. for i in range(len(labels))]

    xopt = run_optim(OSC, df_data, labels, init, lb, ub, id_range, name_prop)
    add_data_to_df(df_data, labels, name_prop, xopt, id_range)


def match_thp(OSC, df_data, id_range):

    name_prop = 'Matching.AVALS[{Hydro2P}][1]'
    labels = [['1', '3'], ['{Manifold}']]
    init = [1. for i in range(len(labels))]
    lb = [0.5 for i in range(len(labels))]
    ub = [5. for i in range(len(labels))]

    set_gap(labels, 'HTCSUR', [df_data.at[0, (label[0].replace('{', '').replace('}', ""), 'htc')] for label in labels])

    xopt = run_optim(OSC, df_data, labels, init, lb, ub, id_range, name_prop)

    add_data_to_df(df_data, labels, name_prop, xopt, id_range)


def run_optim(OSC, df_data, labels, init, lb, ub, id_range, name_prop):

    OSC.set_value('GAP.EnableNetworkValidation', 0)

    xopt = optimization(OSC, df_data, labels, init, lb, ub, name_prop, id_range, 20)

    OSC.set_value('GAP.EnableNetworkValidation', 1)

    return xopt


def optimization(OSC, df_data, labels, init, lb, ub, name_prop, id_range, n_local=20):

    opt = nl.opt(nl.LN_SBPLX, len(init))
    opt.set_lower_bounds(lb)
    opt.set_upper_bounds(ub)
    opt.set_min_objective(lambda x, grad: f(x, grad, OSC, df_data, labels, name_prop, id_range))
    opt.set_maxeval(n_local)

    return opt.optimize(init)


def f(x, grad, OSC, df, labels, name_prop, id_range):

    set_gap(labels, name_prop, x)

    get_from_gap(OSC, df, id_range)

    val = get_value_func(df, name_prop, id_range)

    print(val)

    return val


def set_gap(labels, name_prop, x):
    for idx, label in enumerate(labels):
        for l in label:
            OSC.set_value(f'GAP.MOD[{{PROD}}].PIPE[{l}].{name_prop}', x[idx])


def get_value_func(df, name_prop, id_range):

    if name_prop == 'HTCSUR':
        return 0.5 * sum((df.loc[id_range, ('Separator', 'Temperature')] - df.loc[id_range, ('Separator', 'Temperature Calc')]) ** 2)
    else:
        return 0.5 * (sum((df.loc[id_range, ('W1', 'Tubing Head Pressure')] - df.loc[id_range, ('W1', 'Tubing Head Pressure Calc')]) ** 2) +
                      sum((df.loc[id_range, ('W2', 'Tubing Head Pressure')] - df.loc[id_range, ('W2', 'Tubing Head Pressure Calc')]) ** 2))


def add_data_to_df(df, labels, name_prop, xopt, id_range):

    print(xopt)
    for idx, label in enumerate(labels):
        for l in label:
            if name_prop == 'HTCSUR':
                df.loc[id_range, (l.replace('{', '').replace('}', ""), 'htc')] = xopt[idx]
            else:
                df.loc[id_range, (l.replace('{', '').replace('}', ""), 'fc')] = xopt[idx]


def export_htc_to_excel(df_data, f_path, sh_name):

    sh = xw.Book(f_path).sheets(sh_name)
    sh.range('AD5').options(index=False, header=False).value = df_data.loc[:, ('1', 'htc')]
    sh.range('AF5').options(index=False, header=False).value = df_data.loc[:, ('3', 'htc')]
    sh.range('AH5').options(index=False, header=False).value = df_data.loc[:, ('Manifold', 'htc')]
    xw.Book(f_path).save()


def export_fc_to_excel(df_data, f_path, sh_name):

    sh = xw.Book(f_path).sheets(sh_name)
    sh.range('AE5').options(index=False, header=False).value = df_data.loc[:, ('1', 'fc')]
    sh.range('AG5').options(index=False, header=False).value = df_data.loc[:, ('3', 'fc')]
    sh.range('AI5').options(index=False, header=False).value = df_data.loc[:, ('Manifold', 'fc')]
    xw.Book(f_path).save()


if __name__ == '__main__':

    with OpenServer() as OSC:
        match_temperature_gap_each_time(OSC)
        match_thp_gap_each_time(OSC)

        #match_temperature_gap(OSC)
        #match_thp_gap(OSC)

