#実運用のルールのバックテスト!(excelデータを入力、出力)
#ルール　①Hyperによる取引(日中と夜の立会ごとに決済）　②値幅30円で更新　③前の山(谷)を超えて+30で新規、前の谷(山)を超えて-20で決済　④手数料は売買1回で100円と仮定
import os
import openpyxl as px
from openpyxl.chart import LineChart,Reference,series
import datetime
from collections import deque
from get10min import getSheet,make_performance

#日中立会内での暫定高(安)値とその時の日時を記録する関数
def get225Data():
    
    sheet = getSheet()

    #[]:listを初期化する
    list1 = []
    list2 = []
    list3 = []
    list4 = []
    i = 2
    yori_frag = False
    prev1 = 0
    prev2 = 0
    prev3 = 0
    prev4 = 0

    yori_time_am = datetime.time(8,45,0) 
    hike_time_am = datetime.time(15,15,0)
    yori_time_pm = datetime.time(16,30,0)
    hike_time_pm = datetime.time(5,30,0)

    #cell(1列目)に数値がなくなるまで繰り返す関数
    while True:
        if not sheet.cell(row=i,column=1).value:
            break
        
        #1~3月の場であることを確認
        if 1 <= sheet.cell(row=i,column=1).value.month and sheet.cell(row=i,column=1).value.month <= 3:
            if yori_time_am == sheet.cell(row=i,column=2).value or yori_time_pm == sheet.cell(row=i,column=2).value:
                yori_frag = True

            #立会時間かどうか確認
            if yori_frag == True:
                if not hike_time_am < sheet.cell(row=i,column=2).value and sheet.cell(row=i,column=2).value < yori_time_pm: 
                    #場の開始時に始値と暫定高値の幅が30円(窓開け)の場合
                    if (yori_time_am == sheet.cell(row=i,column=2).value or yori_time_pm == sheet.cell(row=i,column=2).value) and abs(sheet.cell(row=i,column=3).value - prev1) >= 30:
                        #真：始値を暫定高(安)値として、その日時と値段をlistに追加、終値を決める(prev:暫定高(安)値、current:終値)　偽：終値を決める
                        current1 = sheet.cell(row=i,column=3).value
                        n225data = []
                        #list内([0]=日付,[1]=時間,[2]=価格)
                        n225data += [sheet.cell(row=i,column=1).value.date(),sheet.cell(row=i,column=2).value,current1]
                        prev1 = current1
                        current1 = sheet.cell(row=i,column=6).value
                        list1.append(n225data)
                    else:
                        current1 = sheet.cell(row=i,column=6).value

                    #暫定高(安)値と終値の幅が30円以上の場合
                    if abs(current1 - prev1) >= 30:
                        #終値を暫定高(安)値として、日時と値段をlistに追加
                        n225data = []
                        n225data += [sheet.cell(row=i,column=1).value.date(),sheet.cell(row=i,column=2).value,current1]
                        prev1 = current1
                        list1.append(n225data)
                else:
                    yori_frag = False

        elif 3 < sheet.cell(row=i,column=1).value.month and sheet.cell(row=i,column=1).value.month <= 6:
            if yori_time_am == sheet.cell(row=i,column=2).value or yori_time_pm == sheet.cell(row=i,column=2).value:
                yori_frag = True
            
            if yori_frag == True:
                if not hike_time_am < sheet.cell(row=i,column=2).value and sheet.cell(row=i,column=2).value < yori_time_pm:
                    if (yori_time_am == sheet.cell(row=i,column=2).value or yori_time_pm == sheet.cell(row=i,column=2).value) and abs(sheet.cell(row=i,column=3).value - prev1) >= 30:
                        current2 = sheet.cell(row=i,column=3).value
                        n225data = []
                        n225data += [sheet.cell(row=i,column=1).value.date(),sheet.cell(row=i,column=2).value,current2] 
                        prev2 = current2
                        current2 = sheet.cell(row=i,column=6).value
                        list2.append(n225data)
                    else:
                        current2 = sheet.cell(row=i,column=6).value

                    if abs(current2 - prev2) >= 30:
                        n225data = []
                        n225data += [sheet.cell(row=i,column=1).value.date(),sheet.cell(row=i,column=2).value,current2]
                        prev2 = current2
                        list2.append(n225data)
                else:
                    yori_frag = False

        elif 6 < sheet.cell(row=i,column=1).value.month and sheet.cell(row=i,column=1).value.month <= 9:
            if yori_time_am == sheet.cell(row=i,column=2).value or yori_time_pm == sheet.cell(row=i,column=2).value:
                yori_frag = True

            if yori_frag == True:
                if not hike_time_am < sheet.cell(row=i,column=2).value and sheet.cell(row=i,column=2).value < yori_time_pm: 
                    if (yori_time_am == sheet.cell(row=i,column=2).value or yori_time_pm == sheet.cell(row=i,column=2).value) and abs(sheet.cell(row=i,column=3).value - prev1) >= 30:
                        current3 = sheet.cell(row=i,column=3).value
                        n225data = []
                        n225data += [sheet.cell(row=i,column=1).value.date(),sheet.cell(row=i,column=2).value,current3]
                        prev3 = current3
                        current3 = sheet.cell(row=i,column=6).value
                        list3.append(n225data)
                    else:
                        current3 = sheet.cell(row=i,column=6).value

                    if abs(current3 - prev3) >= 30:
                        n225data = []
                        n225data += [sheet.cell(row=i,column=1).value.date(),sheet.cell(row=i,column=2).value,current3]
                        prev3 = current3
                        list3.append(n225data)
                else:
                    yori_frag = False

        elif 9 < sheet.cell(row=i,column=1).value.month and sheet.cell(row=i,column=1).value.month <= 12:
            if yori_time_am == sheet.cell(row=i,column=2).value or yori_time_pm == sheet.cell(row=i,column=2).value:
                yori_frag = True

            if yori_frag == True:
                if not hike_time_am < sheet.cell(row=i,column=2).value and sheet.cell(row=i,column=2).value < yori_time_pm:
                    if (yori_time_am == sheet.cell(row=i,column=2).value or yori_time_pm == sheet.cell(row=i,column=2).value) and abs(sheet.cell(row=i,column=3).value - prev1) >= 30:
                        current4 = sheet.cell(row=i,column=3).value
                        n225data = []
                        n225data += [sheet.cell(row=i,column=1).value.date(),sheet.cell(row=i,column=2).value,current4]
                        prev4 = current4
                        current4 = sheet.cell(row=i,column=6).value
                        list4.append(n225data)
                    else:
                        current4 = sheet.cell(row=i,column=6).value

                    if abs(current4 - prev4) >= 30:
                        n225data = []
                        n225data += [sheet.cell(row=i,column=1).value.date(),sheet.cell(row=i,column=2).value,current4]
                        prev4 = current4
                        list4.append(n225data)
                else:
                    yori_frag = False
        i += 1

    print("データ名を入力してください")
    input_data_kagi = input()
    entry_point1_kagi = []
    entry_point2_kagi = []
    entry_point3_kagi = []
    entry_point4_kagi = []

    kagiSignal(list1,entry_point1_kagi)
    kagiSignal(list2,entry_point2_kagi)
    kagiSignal(list3,entry_point3_kagi)
    kagiSignal(list4,entry_point4_kagi)

    entry_point_kagi = entry_point1_kagi + entry_point2_kagi + entry_point3_kagi + entry_point4_kagi

    make_performance(entry_point_kagi,input_data_kagi+".xlsx")

def kagiSignal(list,entry_point):
    list_count = 0
    retu_count = 1
    #prev_listには暫定高(安)値とトレンド方向を記録,[0]=価格,[1]=トレンド方向
    prev_list = []
    #dequeは先頭、末尾どちらからも取り出せる(キュー、スタック(キューの取り出しはpopleft())
    #yama_tini_queueは直近の山谷形成価格を記録(高値、安値の2つ以上は存在しない)
    yama_tini_queue = deque([])

    #for 変数　in --: --の要素を変数として順番にする
    for w in list:
        #list1の要素1-2で上昇か下降か判断、暫定高(安)値をprev_list1に記録
        if list_count == 0 and list[0][2] > list[1][2]:
            prev_list += [list[1][2],-1]
        elif list_count == 0 and list[0][2] < list[1][2]:
            prev_list += [list[1][2],1]

        #記録されたトレンド方向に記録価格が進んでいたら暫定高(安)値更新
        if list_count >= 2:
            if prev_list[1] == 1 and prev_list[0] < w[2]:
                prev_list[0] = w[2]
                prev_list[1] = 1
            elif prev_list[1] == -1 and prev_list[0] > w[2]:
                prev_list[0] = w[2]
                prev_list[1] = -1
            #トレンドと逆方向に記録価格が進んだ場合
            elif prev_list[1] == 1 and prev_list[0] > w[2]:
                #記録されている直近の山谷形成価格が2つある場合、古い記録を削除し、新しい記録に更新(削除されるのは更新される価格と同じ向き)
                #ない場合は追加
                if len(yama_tini_queue) == 2:
                    yama_tini_queue.popleft()
                    yama_tini_queue.append(prev_list)
                else:
                    yama_tini_queue.append(prev_list)

                prev_list = [w[2],-1]

                retu_count += 1

            elif prev_list[1] == -1 and prev_list[0] < w[2]:
                if len(yama_tini_queue) == 2:
                    yama_tini_queue.popleft()
                    yama_tini_queue.append(prev_list)
                else:
                    yama_tini_queue.append(prev_list)
                
                #prev_list1をStructure()で初期化 ⇒ 初期化しないとyama_tini_queue内の値まで変わるから(Excelで関数コピーった時と同じ動作？)
                prev_list = [w[2],1]

                retu_count += 1

            #yama_tini_queueが2つ記録しているとき、前の記録を取り出し現在価格、記録価格、トレンド方向、価格差、リスト数を表示
            if len(yama_tini_queue) == 2:
                temp = yama_tini_queue.popleft()
                print(w[2],temp[0],temp[1],w[2]-temp[0],list_count)

                #山谷超えのデータをtemp3(価格、日付、シグナル)としてnuki_pointに記録
                #超えなかった場合はyama_tini_queueに戻す
                if temp[1] == 1 and w[2] - temp[0] > 0:
                    temp3 = [w[0],w[1],w[2],"L",retu_count]
                    entry_point.append(temp3)
                elif temp[1] == -1 and temp[0] - w[2] > 0:
                    temp3 = [w[0],w[1],w[2],"S",retu_count]
                    entry_point.append(temp3)
                else:
                    yama_tini_queue.appendleft(temp)
        list_count += 1

get225Data()
