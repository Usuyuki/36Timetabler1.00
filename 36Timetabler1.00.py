import os
import pandas as pd
from time import sleep
import calendar
import datetime
import jpholiday
import datetime

# 起動メッセージ
print("========== ========== ========== ========== ========== ==========")
print("  オンライン授業のためのフォルダ生成ツール(36Timetabler)  ver1.00")
print("========== ========== ========== ========== ========== ==========")
print("Copyright @ 2020 Naofumi Obata. All rights reserved.")
print(" ")
print("ご利用いただき、ありがとうございます。")
sleep(5)

# パスのための変数定義
slash = "/"
period = "."


"""
関数のコーナー

"""
# 定番の終了関数


def End():
    print("20秒後に終了します")
    sleep(20)
    exit()

# 行列変動チェック関数


def Check_sheet(ch_e):
    if str(ch_e) == "qqq":
        print("行及び列に異常操作がないことを確認しました。")
    else:
        print("【行または列が変更されています。申し訳ありませんが正しく処理できません。】")
        End()

# リスト作る関数＋改行と全角スペースを正しく治す関数


def Class_name(now_class_num, now_youbi_num):
    now_class = e_table.iloc[now_class_num, now_youbi_num]
    now_class = str(now_class)
    now_class = now_class.replace("\n", "")
    now_class = now_class.replace("\u3000", "")
    now_class = now_class.replace(" ", "")
    now_class = str(zigen) + period + now_class
    return now_class


"""
ここからプログラム本体

"""

# エクセルファイル読み込み
e_set = pd.read_excel("時間割はここへ(ファイル名変更厳禁！).xlsx", sheet_name="初期設定入力シート")
e_table = pd.read_excel("時間割はここへ(ファイル名変更厳禁！).xlsx", sheet_name="時間割入力シート")


# 行列の変動ないか確認
ch_e_set = e_set.iloc[39, 7]  # y行　列が9を超えるとエラーになる
ch_e_table = e_table.iloc[39, 7]

# 行列変動チェック
print("■Excelシートに行や列の変動がないか確認します。")
sleep(5)
print("初期設定入力シート")
Check_sheet(ch_e_set)
print("時間割入力シート")
Check_sheet(ch_e_table)
sleep(3)

# 初期設定読み込み
"""
0:メインフォルダの生成main_set
1:開始年
2:開始月
3:開始日
4:終了年
5:終了月
6:終了日
7:土日授業フォルダ
8:補習・振替日フォルダ
9:集中講義フォルダ
10:祝日にフォルダ生成するか？yes1no0
"""
main_set = []
for main_set_reading in range(4, 15, 1):
    main_set.append(e_set.iloc[main_set_reading, 1])



"""
初期設定異常確認

"""

print("■初期設定が正しく入力されているか確認します。")
sleep(3)

# 空白ないか？
for nan_check1 in main_set:
    if str(nan_check1) == "nan":
        print("【Excelデータに入力していない項目があるか、Excelが終了されていません。\n全て埋めて、Excelを閉じてから、使用して下さい。】")
        End()
    else:
        print("*")

# 数字で入力されているか？
for z in range(1, 11, 1):
    if type(main_set[z]) == int:
        print("*")
    else:
        print(main_set[z])
        print(type(main_set[z]))
        print("【数字で入力する箇所を文字で入力しています。もう１度Excelデータを確認して下さい】")
        End()

# フォルダ名のstr化
main_set[0] = str(main_set[0])


# 西暦異常
if 2001 <= main_set[1] <= 2200:
    print("*")
else:
    print("【授業開始年が異常です2001年から2200年の間の整数で入力して下さい】")
    End()

if 2001 <= main_set[4] <= 2200:
    print("*")
else:
    print("【授業終了年が異常です2001年から2200年の間の整数で入力して下さい】")
    End()
# 月異常
if 1 <= main_set[2] <= 12:
    print("*")
else:
    print("【授業開始月が異常です1月から12月の間の整数で入力して下さい】")
    End()
if 1 <= main_set[5] <= 12:
    print("*")
else:
    print("【授業終了月が異常です1月から12月の間の整数で入力して下さい】")
    End()
# 日異常
# 参考:https://note.nkmk.me/python-datetime-first-last-date-last-week/
start_end_day_of_month = calendar.monthrange(main_set[1], main_set[2])
end_end_day_of_month = calendar.monthrange(main_set[4], main_set[5])
if 1 <= main_set[3] <= int(start_end_day_of_month[1]):
    print("*")
else:
    print(str(main_set[2])+"月に" + str(main_set[3])+"日は存在しません。正しく入力し直してください。")
    End()
if 1 <= main_set[6] <= int(end_end_day_of_month[1]):
    print("*")
else:
    print(str(main_set[5])+"月に" + str(main_set[6])+"日は存在しません。正しく入力し直してください。")
    End()


dt1 = datetime.datetime(year=main_set[1], month=main_set[2], day=main_set[3],)
dt2 = datetime.datetime(year=main_set[4], month=main_set[5], day=main_set[6],)
dt = dt2 - dt1
dt = dt.days
if dt <= 0:
    print("【授業期間がマイナスです。申し訳ありませんが、過去には戻れません。】")
    End()
else:
    print("*")
if dt >= 365:
    print("【授業期間が長過ぎます。1年以内でお願いします。】")
    End()

else:
    print("*")
if dt <= 31:
    print("【授業期間が短すぎます。31日より長くお願いします。】")
    End()

else:
    print("*")


# 01設定異常
for z in range(7, 11, 1):
    if int(main_set[z]) == 1:
        print("*")
    elif int(main_set[z]) == 0:
        print("*")
    else:
        print("【0か1で入力する所を他の数字で入力しています。もう１度Excelデータを確認して下さい。】")
        End()

print("取得したデータ")
print(main_set)
print("■初期設定を正しく読み込みました。ご入力ありがとうございます。")
sleep(5)


"""
授業読み込み

"""
# 授業読み込み
mon_class = []
tue_class = []
wed_class = []
thi_class = []
fri_class = []
youbi = 0


# 授業をリスト化
for now_youbi_num in range(1, 18, 4):
    youbi += 1
    zigen = 0
    for now_class_num in range(2, 7, 1):
        zigen += 1
        if youbi == 1:
            mon_class.append(Class_name(now_class_num, now_youbi_num))
        elif youbi == 2:
            tue_class.append(Class_name(now_class_num, now_youbi_num))
        elif youbi == 3:
            wed_class.append(Class_name(now_class_num, now_youbi_num))
        elif youbi == 4:
            thi_class.append(Class_name(now_class_num, now_youbi_num))
        elif youbi == 5:
            fri_class.append(Class_name(now_class_num, now_youbi_num))


if len(mon_class) == 0 and len(tue_class) == 0 and len(wed_class) == 0 and len(thi_class) == 0 and len(fri_class) == 0:
    print("【時間割が1つも入力されていないか、Excelデータが終了されていません。今一度お確かめ下さい。】")
    End()
else:
    print("あなたの時間割は")
print("月"+str(mon_class))
print("火"+str(tue_class))
print("水"+str(wed_class))
print("木"+str(thi_class))
print("金"+str(fri_class))

print("※nanと出力されていることがありますが、無記入授業ですので差し支えありません。")
print("■時間割を読み込みました。ご入力ありがとうございます。")
sleep(5)


"""
月・日付リスト生成

"""

# 月リストを生成
print("■フォルダ生成のためのリストを準備しています。しばらくお待ち下さい。")
sleep(4)
mon_hizuke_list = []
tue_hizuke_list = []
wed_hizuke_list = []
thi_hizuke_list = []
fri_hizuke_list = []
start_year = main_set[1]
start_month = main_set[2]
start_day = main_set[3]
end_year = main_set[4]
end_month = main_set[5]
end_day = main_set[6]

# 年越しカウント付き、授業月をリスト化(年を超す場合、翌年分の月もリストに混ぜこむ)
this_month = []
next_month = []
all_month = []
if main_set[1] == main_set[4]:
    for month in range(main_set[2], main_set[5] + 1, 1):
        all_month.append(month)
    this_month = all_month
else:
    print("年越しカウント発動")
    for month in range(main_set[2], 13, 1):
        this_month.append(month)
    x = main_set[5] + 1
    for month in range(1, x, 1):
        next_month.append(month)
    all_month.extend(this_month)
    all_month.extend(next_month)
print(all_month)

"""
日付リスト生成

"""


# 月の最終曜日取得関数　　　　　→参考https://note.nkmk.me/python-datetime-first-last-date-last-week/
def get_day_of_last_week(year, month, dow):  # day of the week =dow = 曜日
    # 月0,火1....日6
    n = calendar.monthrange(year, month)[1]
    l = range(n - 6, n + 1)
    w = calendar.weekday(year, month, l[0])
    w_l = [i % 7 for i in range(w, w + 7)]
    return l[w_l.index(dow)]

# 10以下の数字に0つける関数


def Zeroplus(kazu):
    if kazu < 10:
        kazu = "0" + str(kazu)
    else:
        kazu = str(kazu)
    return kazu


# 祝日除去と日付フォルダ名生成関数
def OPU(year, month, day):
    monthday = "a"
    jholiday = jpholiday.is_holiday(datetime.date(year, month, day))
    if jholiday == True:
        jholiday_name = jpholiday.is_holiday_name(datetime.date(year, month, day))
        if main_set[10] == 1:
            year = year-2000  # 22-0206みたいに年の２文字を前につける
            month = Zeroplus(month)
            day = Zeroplus(day)
            monthday = str(year) + "-" + month + day + "(" + jholiday_name + ")"
        else:
            print(str(jholiday_name) + "なので、この日のフォルダは作りません。充実した休日を過ごして下さいね。")
            monthday = "n"
    else:
        year = year-2000  # 22-0206みたいに年の２文字を前につける
        month = Zeroplus(month)
        day = Zeroplus(day)
        monthday = str(year) + "-" + month + day
    return monthday

# 初月特別処理
print("初月特別処理中")
sleep(1)
month_last = []
for x in range(0, 5, 1):  # 最終日取得
    last = get_day_of_last_week(start_year, start_month, x)
    month_last.append(last)

x = 0
for lday in month_last:  # 曜日ごとにfor
    monthday_list = []
    x += 1
    while int(lday) >= start_day:  # その月が始まるまでfor
        monthday = OPU(start_year, start_month, lday)
        lday -= 7
        if monthday == "n":
            print("*")
        else:
            monthday_list.append(monthday)
    if x == 1:
        mon_hizuke_list.extend(monthday_list)
    elif x == 2:
        tue_hizuke_list.extend(monthday_list)
    elif x == 3:
        wed_hizuke_list.extend(monthday_list)
    elif x == 4:
        thi_hizuke_list.extend(monthday_list)
    elif x == 5:
        fri_hizuke_list.extend(monthday_list)

# 終月特別処理
print("終月特別処理中")
sleep(1)
month_last = []
for x in range(0, 5, 1):
    last = get_day_of_last_week(end_year, end_month, x)
    month_last.append(last)

x = 0
for lday in month_last:  # 曜日ごとにfor
    monthday_list = []
    x += 1
    while int(lday) >= 1:  # その月が始まるまでfor
        monthday = OPU(end_year, end_month, lday)

        if monthday == "n":
            print("*")
        else:
            if int(lday) <= end_day:
                monthday_list.append(monthday)
            else:
                print("*")
        lday -= 7

    if x == 1:
        mon_hizuke_list.extend(monthday_list)
    elif x == 2:
        tue_hizuke_list.extend(monthday_list)
    elif x == 3:
        wed_hizuke_list.extend(monthday_list)
    elif x == 4:
        thi_hizuke_list.extend(monthday_list)
    elif x == 5:
        fri_hizuke_list.extend(monthday_list)

# 中間月処理
if len(all_month) >= 3:  # 3ヶ月未満の授業期間ならこの処理は行わない
    print("中間月処理中")
    sleep(1)
    if len(next_month) == 0:  # 年越さないならe2に終月をぶちこむ
        e2_month = end_month
    else:
        e2_month = 13
        # 翌年処理を先にここで
        print("翌年処理")
        sleep(1)
        for now_month in range(1, end_month, 1):
            month_last = []
            for now_dow in range(0, 5, 1):  # 月の最終曜日作り
                last = get_day_of_last_week(end_year, now_month, now_dow)
                month_last.append(last)
            x = 0
            for lday in month_last:
                monthday_list = []
                x += 1
                while int(lday) >= 1:  # 月初めまで
                    monthday = OPU(end_year, now_month, lday)
                    lday -= 7
                    if monthday == "n":
                        print("*")
                    else:
                        monthday_list.append(monthday)
                if x == 1:
                    mon_hizuke_list.extend(monthday_list)
                elif x == 2:
                    tue_hizuke_list.extend(monthday_list)
                elif x == 3:
                    wed_hizuke_list.extend(monthday_list)
                elif x == 4:
                    thi_hizuke_list.extend(monthday_list)
                elif x == 5:
                    fri_hizuke_list.extend(monthday_list)
    # 初年処理(年越す→12月まで,年越さない→end_monthまで)
    print("初年処理")
    sleep(1)
    e1_month = start_month + 1  # 初月は別処理なので除外
    for now_month in range(e1_month, e2_month, 1):
        month_last = []
        for now_dow in range(0, 5, 1):  # 月の最終曜日作り
            last = get_day_of_last_week(start_year, now_month, now_dow)
            month_last.append(last)
        x = 0
        for lday in month_last:
            monthday_list = []
            x += 1
            while int(lday) >= 1:  # 月初めまで
                monthday = OPU(start_year, now_month, lday)
                lday -= 7
                if monthday == "n":
                    print("*")
                else:
                    monthday_list.append(monthday)
            if x == 1:
                mon_hizuke_list.extend(monthday_list)
            elif x == 2:
                tue_hizuke_list.extend(monthday_list)
            elif x == 3:
                wed_hizuke_list.extend(monthday_list)
            elif x == 4:
                thi_hizuke_list.extend(monthday_list)
            elif x == 5:
                fri_hizuke_list.extend(monthday_list)

else:
    print("*")
# 中間月処理ここまで



print("月曜リスト")
print(mon_hizuke_list)
print("火曜リスト")
print(tue_hizuke_list)
print("水曜リスト")
print(wed_hizuke_list)
print("木曜リスト")
print(thi_hizuke_list)
print("金曜リスト")
print(fri_hizuke_list)


"""
フォルダ作成

"""
# メインフォルダの生成

# もうすでにファイル作っていたら実行しない
maked = os.path.exists(main_set[0])
if maked == True:
    print("【すでに「" + main_set[0] + "」ファイルが生成されています。】")
    End()
else:
    sleep(2)

print("「" + main_set[0] + "」フォルダを生成します。")

# 第１層(タイトル)フォルダの生成
os.mkdir(main_set[0])

# 第２・３層(曜日、授業)生成

# 曜日フォルダとオプションフォルダのパス生成関数
def Option_folder(x, count2, name):
    if x == 1:
        count2 += 1
        print(name+"曜フォルダを生成します。")
        file2_path.append(str(count2) + period + str(name))

    else:
        print(name+"フォルダを生成しません。")
    return count2

# 授業フォルダのパス+フォルダ生成+日付フォルダ生成関数


def Class_fmake(youbi_class, file2, dow):
    for classf in youbi_class:
        if "nan" in str(classf):
            print("空きコマ除去")
        else:
            # 授業フォルダ生成
            path3 = file2 + slash + classf
            os.mkdir(path3)
            print("「"+classf + "」の授業フォルダを生成しました。")

            # 日付フォルダ生成
            count = 0
            if len(youbi_class) != 0:
                if dow == "mon":
                    for a in mon_hizuke_list:
                        path4 = path3 + slash + a
                        os.mkdir(path4)
                        count += 1
                if dow == "tue":
                    for a in tue_hizuke_list:
                        path4 = path3 + slash + a
                        os.mkdir(path4)
                        count += 1
                if dow == "wed":
                    for a in wed_hizuke_list:
                        path4 = path3 + slash + a
                        os.mkdir(path4)
                        count += 1
                if dow == "thi":
                    for a in thi_hizuke_list:
                        path4 = path3 + slash + a
                        os.mkdir(path4)
                        count += 1
                if dow == "fri":
                    for a in fri_hizuke_list:
                        path4 = path3 + slash + a
                        os.mkdir(path4)
                        count += 1
                print(str(count) + "個の日付フォルダを生成しました。")
            else:
                print("*")


# 曜日フォルダ
youbi_name = ["月", "火", "水", "木", "金"]
file2_path = []
count = 0
count2 = 0
for xx in youbi_name:
    count += 1
    count2 = Option_folder(1, count2, xx)


# オプションフォルダ
setting_name = ["土日", "補習・振替日", "集中講義"]
count = 6
count2 = 5
for xx in setting_name:
    count += 1
    count2 = Option_folder(main_set[count], count2, xx)
print("第２層目のフォルダは")
print(file2_path)

# 第２層(曜日)フォルダの生成+第３層(授業)フォルダの生成
count = 0
for path in file2_path:
    file2 = period + slash + main_set[0] + slash + path
    os.mkdir(file2)
    count += 1
    if count == 1:
        Class_fmake(mon_class, file2, "mon")
    elif count == 2:
        Class_fmake(tue_class, file2, "tue")
    elif count == 3:
        Class_fmake(wed_class, file2, "wed")
    elif count == 4:
        Class_fmake(thi_class, file2, "thi")
    elif count == 5:
        Class_fmake(fri_class, file2, "fri")
print("■曜日・オプションフォルダの生成に成功しました。")
sleep(1)
print("■授業フォルダ、日付フォルダの生成に成功しました。")
sleep(2)

# お断り
print("---------------------------------------------------------------------")
if main_set[10] == 1:
    print("学校が定める長期休暇などの日付フォルダも作られています。予めご了承ください。")
else:
    print("日本の祝日を除いた、学校が定める長期休暇などの日付フォルダも作られています。予めご了承ください。")
sleep(7)

# 終了処理
print("---------------------------------------------------------------------")
print("すべての処理が終了しました。お疲れさまでした。")
print("「" + main_set[0] + "」フォルダをお好きな場所へ移動してお使い下さい。")
print("---------------------------------------------------------------------")
sleep(7)
print("exitと入力してEnterキーを押すとこの画面を閉じます。")
aaa = input("ここに入力:")
if str(aaa) == "exit":
    exit()
else:
    sleep(3600)
