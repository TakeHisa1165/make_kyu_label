import xlwings as xw
import sys


class InputToLabel:
    def __init__(self, start_no, no_of_label, path, sheet_name, lot_no):
        """

        :param start_no: 開始位置　GUIで入力した数値
        :param no_of_label: 必要枚数　GUIで入力した数値
        :param path: ラベルファイルのpath
        :param red:　セルの色指定　赤の数値
        :param green:　セルの色指定　緑の数値
        :param blue:　セルの色指定　青の数値
        :param sheet_name:　シート名　path.csvから取得
        :param lot_no:　ロットNo.　GUIで入力した値
        """
        self.path = path
        self.wb = xw.Book(self.path)
        self.ws = self.wb.sheets(sheet_name)
        self.start = start_no
        self.n = no_of_label
        self.cnt = 0
        self.k = 1
        self.ws.range('I5').value = lot_no
        self.ws.range((1, 1), (10000, 6)).color = (255, 255, 255)
        self.ws.range((1, 1), (10000, 6)).clear_contents()
        self.ws.range((1, 20), (10000, 21)).clear_contents()
        self.red = red
        self.green = green
        self.blue = blue
        self.data = 0
        self.input_to_label()

    def input_to_label(self):
        # 偶数列奇数列の判定　偶数の場合
        if self.start % 2 == 0:
            self.data = 1
            self.start_col = 4
            self.row = self.start // 2
            # 何段目のラベルを最初に書き込むか？を判定　1段目
            if self.row == 1:
                self.start_row = 1
                self.ws.range((self.start_row, self.start_col), (self.start_row+4, self.start_col+2)).value = self.ws.range((1, 8), (5, 10)).value
                self.ws.range(self.data, 20).value = self.start_row
                self.ws.range(self.data, 21).value = self.start_col
                self.ws.range((self.start_row, self.start_col), (self.start_row+1, self.start_col + 2)).color = (self.red, self.green, self.blue)
                self.start_row = self.start_row + 5
                self.data += 1
                self.cnt += 1

                self.start_even_col(k=self.k, start_row=self.start_row, cnt=self.cnt, n=self.n, ws=self.ws)
                self.macro()

            # 何段目のラベルを最初に書き込むか？を判定　1段目以外
            else:
                self.cal_row = self.row - 1
                self.start_row = (self.cal_row * 5) + 1
                self.ws.range((self.start_row, self.start_col), (self.start_row+4, self.start_col+2)).value = self.ws.range((1, 8), (5, 10)).value
                self.ws.range(self.data, 20).value = self.start_row
                self.ws.range(self.data, 21).value = self.start_col
                self.ws.range((self.start_row, self.start_col), (self.start_row+1, self.start_col + 2)).color = (self.red, self.green, self.blue)
                self.start_row = self.start_row + 5
                self.data += 1
                self.cnt += 1

                self.start_even_col(k=self.k, start_row=self.start_row, cnt=self.cnt, n=self.n, ws=self.ws)
                self.macro()

        # 偶数列奇数列の判定　奇数の場合
        elif self.start % 2 != 0:
            self.start_col = 1
            self.row = self.start // 2
            # スタート段数の特定　商のみ敬さんするが。0段目はないのでrowが０の場合は１にする
            if self.row == 0:
                self.start_row = 1
                self.start_odd_col(start_row=self.start_row, start_col=self.start_col, cnt=self.cnt, n=self.n,
                                    ws=self.ws)
                self.macro()
            # スタート段数の特定　1段目以外
            else:
                self.cal_row = self.row
                self.start_row = (self.cal_row * 5) + 1
                self.start_odd_col(start_row=self.start_row, start_col=self.start_col, cnt=self.cnt, n=self.n,
                                    ws=self.ws)
                self.macro()

    def start_even_col(self, k, start_row, cnt, n, ws):
        """

        :param k: rowカウント
        :param start_row: 記入始めの行数
        :param cnt: breakまでのカウンター
        :param n: 必要枚数
        :param ws: ワークシート
        :return: なし
        """
        if self.data == 0:
            data = 1
        else:
            data = self.data
        for k in range(start_row, 26, 5):
            for i in range(1, 6, 3):
                ws.range((k, i), (k + 4, i + 2)).value = ws.range((1, 8), (5, 10)).value
                ws.range(data, 20).value = k
                ws.range(data, 21).value = i
                ws.range((k, i), (k + 1, i + 2)).color = (self.red, self.green, self.blue)
                cnt += 1
                data += 1
                if cnt == n:
                    break
            else:
                continue
            break

    def start_odd_col(self, start_row, start_col, cnt, n, ws):
        """

        :param k: rowカウント
        :param start_row: 記入始めの行数
        :param cnt: breakまでのカウンター
        :param n: 必要枚数
        :param ws: ワークシート
        :return: なし
        """
        if self.data == 0:
            data = 1
        else:
            data = self.data
        for k in range(start_row, 26, 5):
            for i in range(start_col, 7, 3):
                ws.range((k, i), (k + 4, i + 2)).value = ws.range((1, 8), (5, 10)).value
                ws.range(data, 20).value = k
                ws.range(data, 21).value = i

                ws.range((k, i), (k + 1, i + 2)).color = (self.red, self.green, self.blue)
                cnt += 1
                data += 1
                if cnt == n:
                    break
            else:
                continue
            break

    def macro(self):
        my_macro = self.wb.macro('Module1.border')
        my_macro()
if __name__ == '__main__':
    pass
    # app = InputToLabel()
    # app.input_to_label()


