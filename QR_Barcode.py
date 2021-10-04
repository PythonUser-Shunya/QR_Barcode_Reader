# このファイルを右上の再生ボタンで実行する
import PySimpleGUI as sg
import function_QR_Barcode


# デザインテーマの設定
sg.theme('DarkTeal7')

# ウィンドウの部品とレイアウト
layout = [
    [sg.Text('読み取り対象のファイルを指定してください')],
    [sg.FileBrowse('①ファイルを選択', file_types=(
        ("Input File", ".xlsx"), )), sg.Input(key='inputFilePath')],
    [sg.Text('②カメラ番号', size=(10, 1)), sg.Combo((0, 1, 2),
                                               default_value=0, size=(10, 1), key='camera_number'), sg.Text("iPoneはどれだ")],
    [sg.Button('③カメラ起動 (Escで終了) ', key='camera')],
    [sg.Text('④場所', size=(10, 1)), sg.Combo(('研究室', '房総', '葉山', '八雲', '水上'),
                                            default_value="研究室", size=(10, 1), key='place'), sg.Text("直接入力も可")],
    [sg.Text('⑤使用者', size=(10, 1)), sg.Combo(('定例', '神田', '松村', '左京楓'),
                                             default_value="", size=(10, 1), key='user'), sg.Text("必要な場合のみ。直接入力も可")],
    [sg.Checkbox('⑥間違いはありませんか？', default=False, key="TF")],
    [sg.Button('⑦保存', key='save', pad=((5, 0), (30, 10)))],
    [sg.Output(size=(80, 20), key="output")]
]

# ウィンドウの生成
window = sg.Window('バーコード読み取り', layout)

# イベントループ
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:  # ウィンドウのXボタンを押したときの処理
        break
    xlsx_file_path = values['inputFilePath']
# QR、バーコード読み込み
    if event == "camera":
        camera_number = values["camera_number"]
        if "電池" in values["inputFilePath"]:
            compiled_list = function_QR_Barcode.make_battery_list()
            read_list = function_QR_Barcode.read_battery(
                camera_number, compiled_list)
            function_QR_Barcode.save_battery(read_list)
            print(f'読み取った電池の総数：{len(read_list)}\n')

        elif "SD" in values["inputFilePath"]:
            compiled_list = function_QR_Barcode.make_SD_list()
            read_list = function_QR_Barcode.read_SD(
                camera_number, compiled_list)
            function_QR_Barcode.save_SD(read_list)
            print(f'読み取ったSDの総数：{len(read_list)}\n')

        else:
            compiled_list = function_QR_Barcode.make_camera_list()
            read_list = function_QR_Barcode.read_camera(
                camera_number, compiled_list)
            function_QR_Barcode.save_camera(read_list)
            print(f'読み取ったカメラの台数：{len(read_list)}\n')

    if event == 'save':
        if values['inputFilePath'] == None:
            print("ファイルを選択してください")
        elif values["TF"] == True:
            place = values["place"]
            if "電池" in values["inputFilePath"]:
                try:
                    ok_make_new_excel_battery = function_QR_Barcode.make_new_excel_battery(
                        xlsx_file_path, place, read_list)
                    if ok_make_new_excel_battery is True:
                        print("保存が完了しました。")
                except PermissionError:
                    print("保存できませんでした。以下が原因である可能性があります。")
                    print("・ウイルスバスターなどのセキュリティーが邪魔をしている", "・Excelファイルを開いている", sep="\n")

            elif "SD" in values["inputFilePath"]:
                if values["user"] is not None:
                    user = values["user"]
                else:
                    user = ""
                try:
                    ok_make_new_excel_SD = function_QR_Barcode.make_new_excel_SD(xlsx_file_path, place, read_list, user)
                    if ok_make_new_excel_SD is True:
                        print("保存が完了しました。")
                except PermissionError:
                    print("保存できませんでした。以下が原因である可能性があります。")
                    print("・ウイルスバスターなどのセキュリティーが邪魔をしている",
                          "・Excelファイルを開いている", sep="\n")
                    
            else:
                if values["user"] is not None:
                    user = values["user"]
                else:
                    user = ""
                try:
                    ok_make_new_excel_camera = function_QR_Barcode.make_new_excel_camera(xlsx_file_path, place, read_list, user)
                    if ok_make_new_excel_camera is True:
                        print("保存が完了しました。")
                except PermissionError:
                    print("保存できませんでした。以下が原因である可能性があります。")
                    print("・ウイルスバスターなどのセキュリティーが邪魔をしている",
                          "・Excelファイルを開いている", sep="\n")
                    
        else:
            print("チェックボックスを確認してください")
window.close()
