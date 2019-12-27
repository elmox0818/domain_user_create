import openpyxl


def edit_excel(values):

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
        sheet['H21'] = values[3]+"@test.com"
        # password
        sheet['H23'] = values[4]

        wb.save('./output/社内連絡票_{}.xlsx'.format(values[8]))
