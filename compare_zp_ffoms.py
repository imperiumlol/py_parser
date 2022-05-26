import traceback

import pandas as pd
import PySimpleGUI as sg

data_mapper_ssch = {
    'cell1':   {
        'row': 'Врачи (кроме зубных), включая врачей-руководителей структурных подразделений',
        'col_ssch': ('Средняя численность работников', 'Работники списочного состава'),
        },
    'cell61':  {
        'row': 'Средний медицинский (фармацев-тический) персонал (персонал, обеспечивающий предоставление медицинских услуг)',
        'col_ssch': ('Средняя численность работников', 'Работники списочного состава'),
        },
    'cell109': {
        'row': 'Младший медицинский (фармацевтический) персонал (персонал, обеспечивающий предоставление медицинских услуг)',
        'col_ssch': ('Средняя численность работников', 'Работники списочного состава'),
        },
    'cell121': {
        'row': 'Руководитель организации',
        'col_ssch': ('Средняя численность работников', 'Работники списочного состава'),
        },
    'cell133': {
        'row': 'Работники, имеющие высшее фармацевтическое или иное высшее образование, предо-ставляющие медицинские услуги (обеспечивающие предоставление медицинских услуг)',
        'col_ssch': ('Средняя численность работников', 'Работники списочного состава'),
        },
    'cell145': {
        'row': 'Прочий персонал',
        'col_ssch': ('Средняя численность работников', 'Работники списочного состава'),
        },
}
data_mapper_fzp = {
    'cell157': {
        'row': 'Врачи (кроме зубных), включая врачей-руководителей структурных подразделений',
        'col_fzp': ('Фонд начисленной заработной платы работников', 'Списочного состава(с внутренним совместительством)'),
        },
    'cell217': {
        'row': 'Средний медицинский (фармацев-тический) персонал (персонал, обеспечивающий предоставление медицинских услуг)',
        'col_fzp': ('Фонд начисленной заработной платы работников', 'Списочного состава(с внутренним совместительством)'),
        },
    'cell265': {
        'row': 'Младший медицинский (фармацевтический) персонал (персонал, обеспечивающий предоставление медицинских услуг)',
        'col_fzp': ('Фонд начисленной заработной платы работников', 'Списочного состава(с внутренним совместительством)'),
        },
    'cell277': {
        'row': 'Руководитель организации',
        'col_fzp': ('Фонд начисленной заработной платы работников', 'Списочного состава(с внутренним совместительством)'),
        },
    'cell289': {
        'row': 'Работники, имеющие высшее фармацевтическое или иное высшее образование, предо-ставляющие медицинские услуги (обеспечивающие предоставление медицинских услуг)',
        'col_fzp': ('Фонд начисленной заработной платы работников', 'Списочного состава(с внутренним совместительством)'),
        },
    'cell301': {
        'row': 'Прочий персонал',
        'col_fzp': ('Фонд начисленной заработной платы работников', 'Списочного состава(с внутренним совместительством)'),
        },
}
data_mapper_fzp_oms = {
     'cell159': {
        'row': 'Врачи (кроме зубных), включая врачей-руководителей структурных подразделений',
        'col_fzp_oms': ('ФЗП списочного состава по источникам финансирования', 'ОМС'),
        },
     'cell219': {
        'row': 'Средний медицинский (фармацев-тический) персонал (персонал, обеспечивающий предоставление медицинских услуг)',
        'col_fzp_oms': ('ФЗП списочного состава по источникам финансирования', 'ОМС'),
        },
     'cell267': {
        'row': 'Младший медицинский (фармацевтический) персонал (персонал, обеспечивающий предоставление медицинских услуг)',
        'col_fzp_oms': ('ФЗП списочного состава по источникам финансирования', 'ОМС'),
        },
     'cell279': {
        'row': 'Руководитель организации',
        'col_fzp_oms': ('ФЗП списочного состава по источникам финансирования', 'ОМС'),
        },
     'cell291': {
        'row': 'Работники, имеющие высшее фармацевтическое или иное высшее образование, предо-ставляющие медицинские услуги (обеспечивающие предоставление медицинских услуг)',
        'col_fzp_oms': ('ФЗП списочного состава по источникам финансирования', 'ОМС'),
        },
     'cell303': {
        'row': 'Прочий персонал',
        'col_fzp_oms': ('ФЗП списочного состава по источникам финансирования', 'ОМС'),
        },
}


def open_excel(path):
    excel_data = pd.read_excel(path, header=[2, 3], index_col=0).round(1)
    excel_data.columns = excel_data.columns.map(lambda x: (x[0].strip(), x[1].strip()))
    return excel_data


def open_json(path):
    json_data = pd.read_json(path).transpose()[0]
    return json_data


def compare(json_data, excel_data, month):
    headings = ['Категория персонала', 'ССЧ в Отраслевом', 'ССЧ в ЗП-ФФОМС', 'Разница' ,'ФЗП в Отраслевом' ,'ФЗП в ЗП-ФФОМС', 'Разница1', 'ФЗП по ОМС в Отраслевом', 'ФЗП по ОМС в ЗП-ФФОМС' , 'Разница2']
    data = []
    for json_key, excel_keys in data_mapper_ssch.items():
##        if excel_keys['row'] not in excel_data.index:
##            continue
        ssch = excel_data.loc[excel_keys['row'], excel_keys['col_ssch']]
        ssch1 = ssch/month
        cell_ = float(json_data[json_key])
        data.append([f"{excel_keys['row']}", round(ssch1, 1), cell_ ,round(ssch1 - cell_, 1)])    
    data_fzp = []
    for idx, (json_key, excel_keys) in enumerate(data_mapper_fzp.items()):
##        if excel_keys['row'] not in excel_data.index:
##            continue
        fzp = excel_data.loc[excel_keys['row'], excel_keys['col_fzp']]
        fzp1 = fzp/1000
        cell_fzp = float(json_data[json_key])
        data[idx].insert(5, round(fzp1, 2))
        data[idx].insert(6, round(cell_fzp, 2))
        data[idx].insert(7, round(fzp1 - cell_fzp, 2))
    data_fzp_oms = []
    for idx, (json_key, excel_keys) in enumerate(data_mapper_fzp_oms.items()):
##        if excel_keys['row'] not in excel_data.index:
##            continue
        fzp_oms = excel_data.loc[excel_keys['row'], excel_keys['col_fzp_oms']]
        fzp_oms1 = fzp_oms/1000
        cell_fzp_oms = float(json_data[json_key])
        data[idx].insert(9, round(fzp_oms1, 2))
        data[idx].insert(10, round(cell_fzp_oms, 2))
        data[idx].insert(11, round(fzp_oms1 - cell_fzp_oms, 2))
    return data, data_fzp, data_fzp_oms, headings



sg.theme("DarkTeal2")
layout = [
    [sg.Text("Открыть Excel-файл: "), sg.Input(key="-EXCEL_PATH-", change_submits=True),
     sg.FileBrowse(key="-EXCEL_PATH-")],
    [sg.Text("Открыть JSON-файл: "), sg.Input(key="-JSON_PATH-", change_submits=True),
     sg.FileBrowse(key="-JSON_PATH-")],
    [sg.Text("Количество месяцев: "), sg.InputText(key="-MONTH-", change_submits=True)],
    [sg.Button("Сравнить")]
]

window = sg.Window('Сравнение ЗП-ФФОМС', layout, size=(550, 150))

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    elif event == "-JSON_PATH-":
        try:
            json_data = open_json(values["-JSON_PATH-"])
        except Exception as e:
            tb = traceback.format_exc()
            sg.popup_error_with_traceback('При открытии файла возникла непредвиденная ошибка!', e, tb)
    elif event == "-EXCEL_PATH-":
        try:
            excel_data = open_excel(values["-EXCEL_PATH-"])
        except Exception as e:
            tb = traceback.format_exc()
            sg.popup_error_with_traceback('При открытии файла возникла непредвиденная ошибка!', e, tb)
    elif event == "Сравнить":
        try:
            data, data_fzp, data_fzp_oms, headings = compare(json_data, excel_data, int(values["-MONTH-"]))
            table_layout = [
                [sg.Table(
                    values=data, headings=headings,
                    auto_size_columns=False,
                    vertical_scroll_only=False,
                    justification='right',
                    #num_rows=1,
                    key='-TABLE-',
                    row_height=35,
                )]
            ]
            sg.Window('Результат сравнения', table_layout).read(close=True)
        except Exception as e:
            tb = traceback.format_exc()
            sg.popup_error_with_traceback('При сравнении возникла непредвиденная ошибка!', e, tb)
