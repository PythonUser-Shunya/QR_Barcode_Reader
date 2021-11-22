# 実行は別ファイルから
# 以下exe化のコマンド
# pyinstaller QR_Barcode.py --onefile --noconsole --add-binary "C:\Users\shuny\anaconda3\envs\barcode\Lib\site-packages\pyzbar\*.dll;pyzbar"
from pyzbar.pyzbar import decode
import cv2
import re
from collections import Counter
import numpy as np
import warnings
import openpyxl
from copy import copy
# import os

# 警告を非表示にする
warnings.simplefilter("ignore")
number_list = []

# 読み込んだ情報を別のファイルに保存する
def save_battery(read_list):
    np.savetxt("battery_list.csv", read_list, delimiter=",", fmt="%s")


def save_SD(read_list):
    np.savetxt("SD_list.csv", read_list, delimiter=",", fmt="%s")


def save_camera(read_list):
    np.savetxt("camera_list.csv", read_list, delimiter=",", fmt="%s")


# ファイルの情報（履歴）を読み込んでリスト型にする。
def load_battery_list():
    file = "battery_list.csv"
    load_list = list(np.loadtxt(
        file, delimiter=",", dtype="unicode", ndmin=1))
    load_list = list(set(load_list))
    return load_list


def load_SD_list():
    file = "SD_list.csv"
    load_list = list(np.loadtxt(
        file, delimiter=",", dtype="unicode", ndmin=1))
    load_list = list(set(load_list))
    return load_list


def load_camera_list():
    file = "camera_list.csv"
    load_list = list(np.loadtxt(
        file, delimiter=",", dtype="unicode", ndmin=1))
    load_list = list(set(load_list))
    return load_list


# もし".csv"があるならその情報（履歴）を使う。なかったら新しくからのリストを作る
def make_battery_list():
    try:
        compiled_list = load_battery_list()
    except OSError:
        compiled_list = []
    return compiled_list


def make_SD_list():
    try:
        compiled_list = load_SD_list()
    except OSError:
        compiled_list = []
    return compiled_list


def make_camera_list():
    try:
        compiled_list = load_camera_list()
    except OSError:
        compiled_list = []
    return compiled_list


# Excelファイルを更新する
def make_new_excel_battery(xlsx_file_path, place, read_list):
    wb = openpyxl.load_workbook(xlsx_file_path)
    ws = wb.worksheets[0]
    for cell_a in ws["A"]:
        if str(cell_a.value) in read_list:
            cell_a_row = cell_a.row
            ws[f"C{cell_a_row}"].value = place
    wb.save(xlsx_file_path)
    return True


def make_new_excel_SD(xlsx_file_path, place, read_list, user):
    wb = openpyxl.load_workbook(xlsx_file_path)
    ws = wb.worksheets[0]
    for cell_a in ws["A"]:
        if cell_a.value in read_list:
            cell_a_row = cell_a.row
            ws[f"E{cell_a_row}"].value = place
            ws[f"F{cell_a_row}"].value = user
    wb.save(xlsx_file_path)
    return True


def make_new_excel_camera(xlsx_file_path, place, read_list, user):
    wb = openpyxl.load_workbook(xlsx_file_path)
    ws = wb.worksheets[0]
    border = copy(ws['A1'].border)
    for cell_a in ws["A"]:
        if str(cell_a.value) in read_list:
            cell_a_row = cell_a.row
            ws[f"B{cell_a_row}"].value = place
            ws[f"C{cell_a_row}"].value = user
            # 罫線
            ws[f"B{cell_a_row}"].border = border
            ws[f"C{cell_a_row}"].border = border
    wb.save(xlsx_file_path)
    return True


# 電池用。30回以上の重複かつ4桁以下かつバーコートのみ許す
def over30_and_under4(number_list, barcode):
    number_list_counter = Counter(number_list)
    number_list_counter_30 = [
        number_list_counter_30[0]
        for
        number_list_counter_30 in number_list_counter.items()
        # if number_list_counter_30[1] >= 30 and len(str(number_list_counter_30[0])) <= 4 and barcode.type == 'CODE128'
        if number_list_counter_30[1] >= 30 and len(str(number_list_counter_30[0])) <= 4
    ]
    return number_list_counter_30

# カメラ用。30回以上の重複かつ4桁以下かつQRのみ許す。
def over30_and_under4(number_list, barcode):
    number_list_counter = Counter(number_list)
    number_list_counter_30 = [
        number_list_counter_30[0]
        for
        number_list_counter_30 in number_list_counter.items()
        if number_list_counter_30[1] >= 30 and len(str(number_list_counter_30[0])) <= 4
        # if number_list_counter_30[1] >= 30 and len(str(number_list_counter_30[0])) <= 4 and barcode.type == 'QRCODE'
    ]
    return number_list_counter_30

# SD用。30回以上の重複かつ5桁以下かつアンダーバーがある情報のみ許す
def over30_and_under5_list(number_list):
    number_list_counter = Counter(number_list)
    number_list_counter_30 = [
        number_list_counter_30[0]
        for
        number_list_counter_30 in number_list_counter.items()
        if number_list_counter_30[1] >= 30 and len(number_list_counter_30[0]) <= 5 and "_" in number_list_counter_30[0]
    ]
    return number_list_counter_30

font = cv2.FONT_HERSHEY_SIMPLEX
# 読み取った総数を表示する関数
def print_total_number(frame, compiled_list):
    cv2.putText(frame, f'COUNT: {len(compiled_list)}', (0, 35),
                font, 1, (0, 0, 0), 3, cv2.LINE_AA)
    cv2.putText(frame, f'COUNT: {len(compiled_list)}', (0, 35),
                font, 1, (255, 255, 255), 1, cv2.LINE_AA)
    cv2.imshow('frame', frame)

# バーコードに四角を描画、読み込んだ情報を表示する関数
def drw_rectangle(frame, barcodeData, x, y, w, h):
    cv2.rectangle(frame, (x, y), (x+w, y+h),(0, 0, 255), 2)
    cv2.rectangle(frame, (x, y-25), (x+60, y), (0, 0, 255), -1)
    cv2.putText(frame, barcodeData, (x, y-10), font, .5, (255, 255, 255), 1, cv2.LINE_AA)

# 電池のバーコードを読み取る
def read_battery(camera_number, compiled_list):
    # 数字のみ読み取れるようにする
    re_compile = re.compile('^[0-9]+$')
    cap = cv2.VideoCapture(int(camera_number))
    while cap.isOpened():
        ret, frame = cap.read()
        if ret == True:
            d = decode(frame)
            if d:
                for barcode in d:
                    # code128のみ許す
                    if barcode.type == 'CODE128':
                        x, y, w, h = barcode.rect
                        barcodeData = barcode.data.decode('utf-8')
                        # 四角を描画
                        drw_rectangle(frame, barcodeData, x, y, w, h)
                        # 読み取った情報を追加する
                        number_list.append(barcodeData)
                        # 読み取った情報にフィルターをかける
                        number_list_counter_30 = over30_and_under4(
                            number_list, barcode)
                        # さらに数字のみしか許さない
                        for c in filter(re_compile.match, number_list_counter_30):
                            compiled_list.append(c)
                            # 重複を許さない
                            compiled_list = list(set(compiled_list))
            # 読み取った総数を表示する
            print_total_number(frame, compiled_list)
        key = cv2.waitKey(10)
        if key == 27:  # ESCkey
            break

    cap.release()
    cv2.destroyAllWindows()
    number_list.clear()
    return compiled_list

# カメラ用
def read_camera(camera_number, compiled_list):
    re_compile = re.compile('^[0-9]+$')
    cap = cv2.VideoCapture(int(camera_number))
    while cap.isOpened():
        ret, frame = cap.read()
        if ret == True:
            d = decode(frame)
            if d:
                for barcode in d:
                    # QRのみ許す
                    if barcode.type == 'QRCODE':
                        x, y, w, h = barcode.rect
                        barcodeData = barcode.data.decode('utf-8')
                        # 四角を描画
                        drw_rectangle(frame, barcodeData, x, y, w, h)
                        # 読み取った情報を追加する
                        number_list.append(barcodeData)
                        # 読み取った情報にフィルターをかける
                        number_list_counter_30 = over30_and_under4(
                            number_list, barcode)
                        # さらに数字のみしか許さない
                        for c in filter(re_compile.match, number_list_counter_30):
                            compiled_list.append(c)
                            # 重複を許さない
                            compiled_list = list(set(compiled_list))
            # 読み取った総数を表示する
            print_total_number(frame, compiled_list)
        key = cv2.waitKey(10)
        if key == 27:  # ESCkey
            break

    cap.release()
    cv2.destroyAllWindows()
    number_list.clear()
    return compiled_list

# SD用
def read_SD(camera_number, compiled_list):
    cap = cv2.VideoCapture(int(camera_number))
    while cap.isOpened():
        ret, frame = cap.read()
        if ret == True:
            d = decode(frame)
            if d:
                for barcode in d:
                    # QRのみ許す
                    if barcode.type == 'QRCODE':
                        x, y, w, h = barcode.rect
                        barcodeData = barcode.data.decode('utf-8')
                        # 四角を描画
                        drw_rectangle(frame, barcodeData, x, y, w, h)
                        # 読み取った情報を追加する
                        number_list.append(barcodeData)
                        # 読み取った情報にフィルターをかける
                        number_list_counter_30 = over30_and_under5_list(
                            number_list)
                        # フィルターを通過した数字のみリストに追加
                        compiled_list += (number_list_counter_30)
                        # 重複を許さない
                        compiled_list = list(set(compiled_list))
            # 読み取った総数を表示する
            print_total_number(frame, compiled_list)
        key = cv2.waitKey(10)
        if key == 27:  # ESCkey
            break

    cap.release()
    cv2.destroyAllWindows()
    number_list.clear()
    return compiled_list
