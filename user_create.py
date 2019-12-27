import PySimpleGUI as sg
import datetime
import lib.sqlserver as sql
import lib.domain as dm
import lib.edit_excel as eex

# 今日の日付取得
today = datetime.date.today().strftime("%Y/%m/%d")
# 所属の選択一覧
dep = ["hogehoge", "fugafuga", "foofoo"]

# 画面レイアウト定義
layout = [
    [sg.Text('ユーザ登録')],
    # 0
    [sg.Text('日付', size=(15, 1)), sg.InputText(today)],
    # 1
    [sg.Text('所属', size=(15, 1)), sg.Combo(default_value="hogehoge",
                                           values=dep, size=(20, 5), enable_events=True)],
    # 2
    [sg.Text('チーム', size=(15, 1)), sg.InputText('')],
    # 3
    [sg.Text('ユーザ名', size=(15, 1)), sg.InputText('')],
    # 4
    [sg.Text('パスワード(Mail)', size=(15, 1)), sg.InputText('')],
    # 5
    [sg.Text('パスワード(PC)', size=(15, 1)), sg.InputText('')],
    # 6
    [sg.Text('メールサーバ', size=(15, 1)), sg.InputText('GMAIL')],
    # 7
    [sg.Text('従業員No', size=(15, 1)), sg.InputText('')],
    # 8
    [sg.Text('氏名', size=(15, 1)), sg.InputText('')],
    # 9
    [sg.Text('よみがな', size=(15, 1)), sg.InputText('')],
    # 10
    [sg.Text('メールアドレス', size=(15, 1)), sg.InputText('@test.com')],
    # 11
    [sg.Text('説明', size=(15, 1)), sg.InputText(today+' test')],
    [sg.Checkbox('データベースへ登録する', default=False),
     sg.Checkbox('ドメインへ登録する', default=False),
     sg.Checkbox('excel出力', default=False)],
    [sg.Submit(button_text='実行ボタン')]
]


window = sg.Window('ユーザ登録ツール').Layout(layout)

while True:
    event, values = window.Read()
    # if event == None or event == 'Exit':
    if event == None:
        print('exit')
        break

    if event == '実行ボタン':
        # DB登録フラグがtrueなら
        # 現在氏名を空白で区切っているか否かのみ
        if len(values[8].split(" ")) > 1:
            if values[12]:
                # sql文作成
                insert_query = "INSERT INTO [hogehoge].[dbo].[foofoo] (dates, belongName, teamName, users, mailPass, pcPass, mailServer, userNo, userName, ruby, mailAddress, domeinDescription, changeDate) VALUES('{0}', '{1}', '{2}', '{3}', '{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}', '{12}')".format(
                    values[0], values[1], values[2], values[3], values[4], values[5], values[6], values[7], values[8], values[9], values[10], values[11], today)
                # sql文実行
                sql.ExecuteQueryBySQLServer(insert_query)
            if values[13]:
                userInfo = {
                    'belongName': values[2],
                    'teamName': values[3],
                    # 一覧の表示名
                    'cn': values[3],
                    # 姓
                    'sn': values[8].split(" ")[0],
                    # 表示名
                    'displayName': values[8],
                    # ログインID
                    'sAMAccountName': values[3],
                    # 名
                    'givenName': values[8].split(" ")[1],
                    # ユーザログオン名
                    'userPrincipalName': values[3],
                    # 説明
                    'description': values[11],
                }
                # ドメイン登録
                dm.user_add(userInfo)
            if values[14]:
                # Excel出力
                eex.edit_excel(values)

            show_message = "完了！"

        else:
            show_message = "エラー 入力値を確認してください"

        # ポップアップ
        sg.Popup(show_message)
