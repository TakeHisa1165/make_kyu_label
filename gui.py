import PySimpleGUI as sg
import xlwings as xw
import os
import csv
import w_csv
import sys
import win32com.client
import input_label


class InputWindow:
    def __init__(self):
        #　ラベル作成エクセルブックを開く
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        try:
            with open('path.csv', newline='') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    self.dir_path = row["dir_path"]
                    self.sheet_name = row["sheet_name"]

        except FileNotFoundError:
            sg.popup_ok('初期設定が必要です。\n設定画面から書き出しフォルダを設定してください。')
            SelectFile()

        try:
            self.label_file_path = self.dir_path
            xw.Book(self.label_file_path)
        except AttributeError:
            sg.popup_error('Excelファイルを開けません\n初期設定をやり直してください。')
            SelectFile()

    def input_window(self):
        sg.theme("systemdefault")

        layout = [
            [sg.MenuBar([["設定",["基本設定"]]], key="menu1")],
            [sg.Text("開始位置", font=("メイリオ", 14)), sg.InputText(size=(5, 1), key="-start_no-", font=("メイリオ", 14)),
                sg.Text("必要数", font=("メイリオ", 14)), sg.InputText(size=(5, 1), key="-no_of_label-", font=("メイリオ", 14))],
            [sg.Text("ロット番号", font=("メイリオ", 14)), sg.InputText(size=(17, 1), key="-lot_no-", font=("メイリオ", 14))],
            [sg.Submit(button_text="ラベル作成", size=(10, 1), pad=((50, 100), (0, 0)),font=("メイリオ", 14)), sg.Submit(button_text="終了", font=("メイリオ", 14))],
            ]


        window = sg.Window('ラベル作成', layout)

        while True:
            event, values = window.read()

            if event is None:
                print(exit)
                sys.exit()

            if event == "ラベル作成":
                start_no = values["-start_no-"]
                start_no = int(start_no)
                no_of_label = values["-no_of_label-"]
                no_of_label = int(no_of_label)
                lot_no = values["-lot_no-"]
                path = self.label_file_path
                input_label.InputToLabel(start_no=start_no, no_of_label=no_of_label, path=path,
                                        sheet_name=self.sheet_name, lot_no = lot_no)


            if event == "終了":
                sys.exit()
            # if event == "印刷":
            #     input_to_excel.PrintOut(self.label_file_path)

            if values["menu1"] == "基本設定":
                SelectFile()


        window.close()

class SelectFile:
    def __init__(self):
        self.path_dict = self.select_file()

    def select_file(self):

        sg.theme("systemdefault")


        layout = [
            [sg.Text("ラベル作成ファイルを選んでください", size=(50, 1), font=('メイリオ', 14))],
            [sg.InputText(font=('メイリオ', 14), key="-dir_path-"), sg.FileBrowse('開く', key='File1', font=('メイリオ', 14))],
            [sg.Text("シート名を入力してください", size=(50, 1), font=('メイリオ', 14))],
            [sg.InputText(font=('メイリオ', 14), key="-sheet_name-")],
            [sg.Submit(button_text='設定', font=('メイリオ', 14)), sg.Submit(button_text="閉じる", font=('メイリオ', 14))],

        ]

        # セクション 2 - ウィンドウの生成z
        window = sg.Window('ファイル選択', layout)

        # セクション 3 - イベントループ
        while True:
            event, values = window.read()

            if event is None:
                print('exit')
                sys.exit()

            if event == '設定':
                path_dict = {}
                dir_path = values["-dir_path-"]
                path_dict["dir_path"] = dir_path
                sheet_name = values["-sheet_name-"]
                path_dict["sheet_name"] = sheet_name
                csv = w_csv.Write_csv()
                csv.write_csv(path_dict=path_dict)
                sg.popup('初期設定が完了しましたアプリを再起動してください\nアプリを終了します')
                sys.exit()


                return path_dict
            if event == '閉じる':
                sys.exit()




        #  セクション 4 - ウィンドウの破棄と終了
        window.close()

if __name__ == "__main__":
    app = InputWindow()
    app.input_window()