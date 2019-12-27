import openpyxl


def edit_excel(values):

    #    # 0
    #    [sg.Text('日付', size=(15, 1)), sg.InputText(today)],
    #    # 1
    #    [sg.Text('所属', size=(15, 1)), sg.Combo(default_value="トラクタ技術部",
    #                                           values=dep, size=(20, 5), enable_events=True)],
    #    # 2
    #    [sg.Text('チーム', size=(15, 1)), sg.InputText('')],
    #    # 3
    #    [sg.Text('ユーザ名', size=(15, 1)), sg.InputText('')],
    #    # 4
    #    [sg.Text('パスワード(Mail)', size=(15, 1)), sg.InputText('')],
    #    # 5
    #    [sg.Text('パスワード(PC)', size=(15, 1)), sg.InputText('')],
    #    # 6
    #    [sg.Text('メールサーバ', size=(15, 1)), sg.InputText('GMAIL')],
    #    # 7
    #    [sg.Text('従業員No', size=(15, 1)), sg.InputText('')],
    #    # 8
    #    [sg.Text('氏名', size=(15, 1)), sg.InputText('')],
    #    # 9
    #    [sg.Text('よみがな', size=(15, 1)), sg.InputText('')],
    #    # 10
    #    [sg.Text('メールアドレス', size=(15, 1)), sg.InputText('@kubota.com')],
    #    # 11
    #    [sg.Text('説明', size=(15, 1)), sg.InputText(today+' test')],

    if len(values[4]) < 1:
        print("pc_template")
        wb = openpyxl.load_workbook('./template_pc_pass.xlsx')
        sheet = wb['Sheet1']
        # 部
        sheet['E6'] = values[1]
        # 名前
        sheet['E8'] = values[8]
        # 日付
        sheet['W4'] = values[0]
        # username
        sheet['H16'] = values[3]
        # password
        sheet['H18'] = values[5]

        wb.save('./output/社内連絡票_{}.xlsx'.format(values[8]))

    else:
        print("mail_template")
        wb = openpyxl.load_workbook('./template_pc_mail_pass.xlsx')
        sheet = wb['Sheet1']

        # 部
        sheet['E6'] = values[1]
        # 名前
        sheet['E8'] = values[8]
        # 日付
        sheet['W4'] = values[0]
        # username
        sheet['H16'] = values[3]
        # password
        sheet['H18'] = values[5]
        # username
        sheet['H21'] = values[3]+"@kubota.com"
        # password
        sheet['H23'] = values[4]

        wb.save('./output/社内連絡票_{}.xlsx'.format(values[8]))
