import random
import datetime
import os
import sys
import pandas as pd
import PySimpleGUI as sg
import configparser
from styleframe import StyleFrame

pd.options.mode.chained_assignment = None
sg.ChangeLookAndFeel('GreenTan')

hint_member = """
說明：
1. 一行即代表一個成員。程式會從清單內隨機抽取，已中獎者會挪至已中獎清單內。
2. 單行以逗號(,)分隔，最開頭必代表編號。編號不可重複。
3. 每一行格式需相同。

格式：
1.「編號」。僅為正整數，且編號不可重複。例如：1 <-- 表示編號1
2.「編號,人名」。例如：2,王小明
3.「編號,工號,人名」。例如：2,111999,王小明
4.「編號,工號,部門名稱,人名」。例如：2,111999,部門A,王小明

若用xlsx或csv檔案上傳，則限定資料欄位數量不超過3個。
並依照欄位數量與名稱對應上述第2~4的格式，程式在上傳後會自動給定編號：
2.「人名」
3.「工號」、「人名」
4.「工號」、「部門名稱」、「人名」
"""

hint_prize = """
說明：
1. 一行即代表一個獎項，越上面越先抽。抽完的獎項會挪至已抽獎項清單內。
2. 單行以逗號(,)分隔，最末尾必代表獎品數量。
3. 每一行格式需相同。

格式：
1.「獎品數量」。僅有數字，表示產品數量。例如：10 <-- 表示抽10個
2.「獎項,獎品數量」。例如：三獎,5 <-- 表示三獎有5個名額
3.「獎項,獎品內容,獎品數量」。例如：三獎,15000元,5 <-- 表示三獎為15000元，且有5個名額

若用xlsx或csv檔案上傳，則限定資料欄位數量不超過3個，並依照欄位數量與名稱對應上述3種格式：
1.「獎品數量」
2.「獎項」、「獎品數量」
3.「獎項」、「獎品內容」、「獎品數量」
"""

hint_savefile = """
若設定抽獎模式為一次全開，則會直接產出檔案；若設定逐次開獎，則會逐獎項產出中獎紀錄的檔案。
如果產出檔案為無，則不會有動作。
"""

hint_sendmail = """
依照信件收件人清單，在每次抽獎後都會以excel的方式寄送出去，並以抽獎獎項為檔案名稱。
"""

config = configparser.ConfigParser()
config.read('config.ini', encoding = 'utf-8')
savefilepath = config['DEFAULT']['savefilepath'] if len(config['DEFAULT']['savefilepath']) > 0 else None
mode_member, mode_prize = 0, 0
count = 0
Mline_member_size = (35, 8)
Mline_prize_size = (25, 8)
Mline_current_prize_width = 50
Mline_result_width = 55
dict_member_01, list_prize_01 = {}, []
dict_member_02, list_prize_02 = {}, []
dict_record = {}
extr_num, remain_num = 0, 0
current_prize = None

layout_input = [
    [sg.Text('抽獎程式', justification = 'center', font = ('Microsoft YaHei', 15), relief = sg.RELIEF_RIDGE)],
    [
        sg.Column([
            [
                sg.Column([
                        [sg.Text('待抽成員名單', size = (20, 1), auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10))],
                        [sg.Multiline("", key = 'Mline_member_01', size = Mline_member_size, expand_x = True, expand_y = True, tooltip = hint_member)]
                ]),
                sg.Column([
                        [sg.Text('已抽成員名單', size = (20, 1), auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10))],
                        [sg.Multiline("", key = 'Mline_member_02', size = Mline_member_size, expand_x = True, expand_y = True, tooltip = hint_member)]
                ])
            ],
            [
                sg.FileBrowse('載入成員資料', target = 'filename_member',  font = ('Microsoft YaHei', 10), file_types = (('xlsx', '*.xlsx'), ('csv', '*.csv'), ('All Files', '*.*'))),
                sg.Button(button_text = '上傳', key = 'upload_member', font = ('Microsoft YaHei', 10)),
                sg.Button(button_text = '清空', key = 'clean_member', font = ('Microsoft YaHei', 10))
            ]
        ]),
        sg.VSeparator(),
        sg.Column([
            [
                sg.Column([
                        [sg.Text('待抽獎項清單', size = (20, 1), auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10))],
                        [sg.Multiline("", key = 'Mline_prize_01', size = Mline_prize_size, expand_x = True, expand_y = True, tooltip = hint_prize)]
                ]),
                sg.Column([
                        [sg.Text('已抽獎項清單', size = (20, 1), auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10))],
                        [sg.Multiline("", key = 'Mline_prize_02', size = Mline_prize_size, expand_x = True, expand_y = True, tooltip = hint_prize)]
                ])
            ],
            [
                sg.FileBrowse('載入獎項資料', target = 'filename_prize',  font = ('Microsoft YaHei', 10), file_types = (('xlsx', '*.xlsx'), ('csv', '*.csv'), ('All Files', '*.*'))),
                sg.Button(button_text = '上傳', key = 'upload_prize', font = ('Microsoft YaHei', 10)),
                sg.Button(button_text = '清空', key = 'clean_prize', font = ('Microsoft YaHei', 10))
            ]
        ])
    ],
    [sg.HSeparator()],
    [
        sg.Column([
            [
                sg.Text('抽獎模式', size = (7, 1), auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10)),
                sg.Combo(['逐次開獎', '一次全開！！'], default_value = '逐次開獎', key = 'lottery_mode', size = (15, 40), enable_events = True, font = ('Microsoft YaHei', 10))
            ],
            [sg.Checkbox('在每次抽獎後跳出視窗', key = 'bool_popup_result', font = ('Microsoft YaHei', 10), disabled = True)],
            [
                sg.Text('輸出檔案', size = (7, 1), auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10), tooltip = hint_savefile),
                sg.Combo(['無', 'txt檔', 'xlsx檔'], default_value = '無', key = 'output_mode', size = (15, 40), enable_events = True, font = ('Microsoft YaHei', 10), tooltip = hint_savefile)
            ],
            [sg.Checkbox('在每次抽獎後寄出該次結果', key = 'bool_sendmail', font = ('Microsoft YaHei', 10), tooltip = hint_sendmail, disabled = True)],
            [
                sg.Column([
                    [sg.Button(button_text = '清空中獎紀錄', key = 'clean_record', font = ('Microsoft YaHei', 10))],
                    [sg.Button(button_text = '清空全部資料', key = 'clean_all', font = ('Microsoft YaHei', 10))]
                ]),
                sg.Button(button_text = '準備', key = 'Do', font = ('Microsoft YaHei', 24), button_color = '#475841')# button_color = '#475841' (default)
            ]
        ]),
        sg.Column([
            [sg.Text('逐次抽獎獎項：', size = (Mline_current_prize_width - 5, 1), key = 'current_prize_str', auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10)),],
            [   
                sg.Text('抽取數量', size = (7, 1), auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10)),
                sg.InputText(
                    '', 
                    key = 'draw_num', 
                    size = (4, 1), 
                    font = ('Microsoft YaHei', 10), 
                    use_readonly_for_disable = False,  # 有readonly跟disalbed兩種參數不能共存的要選擇
                    disabled = False,
                ),
                sg.Text('已抽數量：', size = (7 + 6, 1), key = 'extr_str', auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10)),
                sg.Text('剩餘數量：', size = (7 + 6, 1), key = 'remain_str', auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10)),
            ],
            [sg.Text('本次中獎名單', size = (20, 1), auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10))],
            [sg.Multiline("", key = 'Mline_result', size = (Mline_current_prize_width, 8), expand_x = True, expand_y = True,)]
        ]),
        sg.Column([
            [sg.Text('中獎紀錄', size = (20, 1), auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 10))],
            [sg.Multiline("", key = 'Mline_allrecord', size = (Mline_result_width, 12), expand_x = True, expand_y = True,)]
        ]),
    ],
    [sg.HSeparator()],
    [
        [
            sg.Text('成員清單路徑:', auto_size_text = True, font = ('Microsoft YaHei', 9)),
            sg.Text(size = (100, 1), key = 'filename_member', auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 9)) # relief = sg.RELIEF_RIDGE 最後可拿掉
        ],
        [
            sg.Text('獎品清單路徑:', auto_size_text = True, font = ('Microsoft YaHei', 9)),
            sg.Text(size = (100, 1), key = 'filename_prize', auto_size_text = True, justification = 'left', font = ('Microsoft YaHei', 9)) # relief = sg.RELIEF_RIDGE 最後可拿掉
        ]   
    ]
]

window_input = sg.Window(
    '抽獎程式', 
    layout_input, 
    default_element_size = (80, 1), 
    grab_anywhere = False,
    resizable = True,
)

while True:  
    event, input_values = window_input.read()

    if event == sg.WIN_CLOSED: # event == '取消'
        sg.PopupOK('作業取消', font = ('Microsoft YaHei', 10))
        break

    if event == 'clean_member':  # 清空成員名單與成員資料路徑
        window_input['Mline_member_01'].Update('')
        window_input['Mline_member_02'].Update('')
        window_input['filename_member'].Update('')
        dict_member_02 = {}

    if event == 'clean_prize':  # 清空獎項清單與獎項資料路徑
        window_input['Mline_prize_01'].Update('')
        window_input['Mline_prize_02'].Update('')
        window_input['filename_prize'].Update('')
        count = 0
        list_prize_02 = []

    if event == 'clean_record':  # 清空當次結果
        window_input['Mline_result'].Update('')
        window_input['Mline_allrecord'].Update('')
        window_input['current_prize_str'].update('逐次抽獎獎項：')
        window_input['draw_num'].update('')
        window_input['extr_str'].update('已抽數量：')
        window_input['remain_str'].update('剩餘數量：')
        count = 0
        extr_num, remain_num = 0, 0
        dict_record = {}
    
    if event == 'clean_all':  # 全清空
        window_input['Mline_member_01'].Update('')
        window_input['Mline_member_02'].Update('')
        window_input['Mline_prize_01'].Update('')
        window_input['Mline_prize_02'].Update('')
        window_input['filename_member'].Update('')
        window_input['filename_prize'].Update('')
        window_input['Mline_result'].Update('')
        window_input['Mline_allrecord'].Update('')
        window_input['current_prize_str'].update('逐次抽獎獎項：')
        window_input['draw_num'].update('')
        window_input['extr_str'].update('已抽數量：')
        window_input['remain_str'].update('剩餘數量：')
        input_values['載入成員資料'] = ''
        count = 0
        extr_num, remain_num = 0, 0
        dict_member_02, list_prize_02 = {}, []
        dict_record = {}
    
    if event == 'upload_member': # 上傳成員檔案
        if len(input_values['載入成員資料']) == 0:
            sg.PopupOK('尚未載入成員資料。', font = ('Microsoft YaHei', 10))
        else:
            try:
                file_path_member = input_values['載入成員資料']
                Data_upload_member = pd.read_excel(file_path_member)

                if Data_upload_member.columns.tolist() == ['工號', '部門名稱', '人名']:
                    mode_member = 3
                    Data_upload_member['工號'] = Data_upload_member['工號'].astype(str)
                    Data_upload_member['編號'] = Data_upload_member.index + 1
                    Data_upload_member['main'] = Data_upload_member.apply(lambda x: ','.join(x[['編號', '工號', '部門名稱', '人名']].astype(str)), axis = 1)
                    window_input['Mline_member_01'].Update('\n'.join(Data_upload_member['main']))
                elif Data_upload_member.columns.tolist() == ['工號', '人名']:
                    mode_member = 2
                    Data_upload_member['工號'] = Data_upload_member['工號'].astype(str)
                    Data_upload_member['編號'] = Data_upload_member.index + 1
                    Data_upload_member['main'] = Data_upload_member.apply(lambda x: ','.join(x[['編號', '工號', '人名']].astype(str)), axis = 1)
                    window_input['Mline_member_01'].Update('\n'.join(Data_upload_member['main']))
                elif Data_upload_member.columns.tolist() == ['人名']:
                    mode_member = 1
                    Data_upload_member['編號'] = Data_upload_member.index + 1
                    Data_upload_member['main'] = Data_upload_member.apply(lambda x: ','.join(x[['編號', '人名']].astype(str)), axis = 1)
                    window_input['Mline_member_01'].Update('\n'.join(Data_upload_member['main']))
                else:
                    sg.PopupOK('載入成員資料不符合格式。', font = ('Microsoft YaHei', 10))
            except:
                sg.PopupOK('載入成員資料失敗，請確認。', font = ('Microsoft YaHei', 10))
    
    if event == 'upload_prize': # 上傳獎項檔案
        if len(input_values['載入獎項資料']) == 0:
            sg.PopupOK('尚未載入獎項資料。', font = ('Microsoft YaHei', 10))
        else:
            try:
                file_path_prize = input_values['載入獎項資料']
                Data_upload_prize = pd.read_excel(file_path_prize)

                if Data_upload_prize.columns.tolist() == ['獎項', '獎品內容', '獎品數量']:
                    mode_prize = 3
                    Data_upload_prize['獎品數量'] = Data_upload_prize['獎品數量'].astype(int)
                    Data_upload_prize['編號'] = Data_upload_prize.index + 1
                    Data_upload_prize['main'] = Data_upload_prize.apply(lambda x: ','.join(x[['獎項', '獎品內容', '獎品數量']].astype(str)), axis = 1)
                    window_input['Mline_prize_01'].Update('\n'.join(Data_upload_prize['main']))
                elif Data_upload_prize.columns.tolist() == ['獎項', '獎品數量']:
                    mode_prize = 2
                    Data_upload_prize['獎品數量'] = Data_upload_prize['獎品數量'].astype(int)
                    Data_upload_prize['編號'] = Data_upload_prize.index + 1
                    Data_upload_prize['main'] = Data_upload_prize.apply(lambda x: ','.join(x[['獎項', '獎品數量']].astype(str)), axis = 1)
                    window_input['Mline_prize_01'].Update('\n'.join(Data_upload_prize['main']))
                elif Data_upload_prize.columns.tolist() == ['獎品數量']:
                    mode_prize = 1
                    Data_upload_prize['獎品數量'] = Data_upload_prize['獎品數量'].astype(int)
                    Data_upload_prize['編號'] = Data_upload_prize.index + 1
                    Data_upload_prize['main'] = Data_upload_prize['獎品數量'].astype(str)
                    window_input['Mline_prize_01'].Update('\n'.join(Data_upload_prize['main']))
                else:
                    sg.PopupOK('載入獎項資料不符合格式。', font = ('Microsoft YaHei', 10))
            except:
                sg.PopupOK('載入獎項資料失敗，請確認。', font = ('Microsoft YaHei', 10))

    if event == 'lottery_mode': # 切換抽獎模式(抽獎功能初始化)
        if input_values['lottery_mode'] == '逐次開獎':
            window_input['Do'].update('準備', button_color = '#475841')
            window_input['draw_num'].update(disabled = False)
        elif input_values['lottery_mode'] == '一次全開！！':
            window_input['Do'].update('抽獎', button_color = 'red')
            window_input['draw_num'].update(disabled = True)
        else:
            sg.PopupOK('怎麼有其他抽獎模式？', font = ('Microsoft YaHei', 10))

    if event == 'Do':  # 抽獎按鈕
        # ####測試區 跳出中獎大名單####
        # if input_values['bool_popup_result']:
        #     layout_popup_winner = [
        #         [sg.Text('中獎名單', justification = 'center', font = ('Microsoft YaHei', 15), relief = sg.RELIEF_RIDGE)],
        #         [sg.Button(button_text = '確定', key = 'ok', font = ('Microsoft YaHei', 10))]
        #     ]

        #     window_popup_winner = sg.Window(
        #         '中獎名單', 
        #         layout_popup_winner, 
        #         size = (200, 200), 
        #         # grab_anywhere = False,
        #         # resizable = True,
        #     )

        #     while True:  
        #         event_popup, input_values_popup = window_popup_winner.read()
        #         if event_popup in [sg.WIN_CLOSED, 'ok']:
        #             break
        #     window_popup_winner.close()

        # ####測試區####


        if input_values['output_mode'] in ['txt檔', 'xlsx檔']:
            try:
                if not os.path.exists(savefilepath):
                    os.mkdir(savefilepath)
            except:
                sg.PopupOK('請確認輸出路徑。', font = ('Microsoft YaHei', 10))

        # 上傳跟實際抽獎是兩回事，以待抽清單內的資料為準，所以mode_member & mode_prize也都要重新判斷
        if len(input_values['Mline_member_01']) == 0: # 待抽成員為空白
            sg.PopupOK('無待抽成員。', font = ('Microsoft YaHei', 10))
        elif len(input_values['Mline_prize_01']) == 0: # 待抽獎項為空白
            sg.PopupOK('無待抽獎項。', font = ('Microsoft YaHei', 10))
        elif len(set([i.count(',') for i in input_values['Mline_member_01'].split('\n')])) > 1:  # 待抽成員格式不統一
            sg.PopupOK('請確認待抽成員格式是否統一。', font = ('Microsoft YaHei', 10))
        elif len(set([i.count(',') for i in input_values['Mline_prize_01'].split('\n')])) > 1: # 待抽獎項格式不統一
            sg.PopupOK('請確認待抽獎項格式是否統一。', font = ('Microsoft YaHei', 10))
        elif list(set([i.count(',') for i in input_values['Mline_member_01'].split('\n')]))[0] not in {0, 1, 2, 3}: # 待抽成員要符合格式(mode_member)
            sg.PopupOK('請確認待抽成員名單是否符合格式。', font = ('Microsoft YaHei', 10))
        elif list(set([i.count(',') for i in input_values['Mline_prize_01'].split('\n')]))[0] not in {0, 1, 2}: # 待抽獎項要符合格式(mode_prize)
            sg.PopupOK('請確認待抽獎項清單是否符合格式。', font = ('Microsoft YaHei', 10))
        else:
            mode_member = list(set([i.count(',') for i in input_values['Mline_member_01'].split('\n')]))[0]
            mode_prize = list(set([i.count(',') for i in input_values['Mline_prize_01'].split('\n')]))[0]

            # 待抽成員名單，整理成dict {編號:對象, ...}
            # 若名單只有編號(mode_member = 0)，則編號本身亦為value
            # 否則(mode_member in {1, 2, 3}) value為每一行第一個逗點(,)之後的資料
            if mode_member == 0: 
                dict_member_01 = {i:i for i in input_values['Mline_member_01'].split('\n')}
                dict_member_02 = {i:i for i in input_values['Mline_member_02'].split('\n')}
            else:
                dict_member_01 = {i.split(',', 1)[0]:i.split(',', 1)[-1] for i in input_values['Mline_member_01'].split('\n')}
                dict_member_02 = {i.split(',', 1)[0]:i.split(',', 1)[-1] for i in input_values['Mline_member_02'].split('\n')}
            
            # 待抽獎項清單，整理成list [[項目, 獎品數量], ...]
            # 若清單只有獎品數量(mode_prize = 0)，則list中的list只放獎品數量，可用count代替為項目
            # 否則(mode_member in {1, 2}) list中的list會放項目與獎品數量
            list_prize_01 = [i.rsplit(',', 1) for i in input_values['Mline_prize_01'].split('\n')]
            list_prize_02 = [i.rsplit(',', 1) for i in input_values['Mline_prize_02'].split('\n')]

            # print(f'待抽成員名單：\n{dict_member_01}')
            # print(f'已抽成員名單：\n{dict_member_02}')
            # print(f'待抽獎項清單：\n{list_prize_01}')
            # print(f'已抽獎項清單：\n{list_prize_02}')

            try:
                # current_prize = None
                if len(input_values['Mline_member_01'].split('\n')) > len(dict_member_01): # 表示待抽成員名單有重複編號
                    sg.PopupOK('待抽成員名單編號有重複，無法執行抽獎。', font = ('Microsoft YaHei', 10))
                elif len(dict_member_01) < sum([int(i[-1]) for i in list_prize_01]): # 總成員數量小於總獎項數量
                    sg.PopupOK('待抽成員名單數量小於待抽獎項數量，無法執行抽獎。', font = ('Microsoft YaHei', 10))
                else:
                    if input_values['lottery_mode'] == '逐次開獎':
                        str_original_record = input_values['Mline_allrecord'] # 先存取當下的開獎紀錄

                        if window_input['Do'].get_text() == '準備':  # 逐次抽獎準備階段
                            if remain_num == 0: # 如果全部抽完(剩餘0)則直接拉下一個獎項
                                current_prize = list_prize_01[0] # 當次獎項與數量
                                k_num = int(current_prize[-1])
                                current_prize_str = f'第 {count + 1} 個獎項' if mode_prize == 0 else ','.join(current_prize[:-1])
                                window_input['current_prize_str'].update('逐次抽獎獎項：{}'.format(current_prize_str))
                                window_input['draw_num'].update('{}'.format(k_num)) # 預設全抽
                                window_input['extr_str'].update('已抽數量：{}'.format(0))
                                window_input['remain_str'].update('剩餘數量：{}'.format(k_num))
                                extr_num = int(window_input['extr_str'].DisplayText.split('：')[-1])
                                remain_num = int(window_input['remain_str'].DisplayText.split('：')[-1])

                            window_input['Do'].update('抽獎', button_color = 'red')
                        elif window_input['Do'].get_text() == '抽獎':
                            try:
                                print('逐次開獎')
                                k_num = int(input_values['draw_num'])
                            except:
                                sg.PopupOK('抽取數量請輸入數字。', font = ('Microsoft YaHei', 10))
                            else:
                                if k_num > remain_num:
                                    sg.PopupOK('抽取數量不可大於剩餘數量，將預設為剩餘數量。', font = ('Microsoft YaHei', 10))
                                    window_input['draw_num'].update('{}'.format(remain_num)) # 預設剩下的最大數目
                                else:

                                    # ###################### 逐次開獎 Start ######################
                                    dict_result = {}
                                    dict_member_02 = dict(reversed(list(dict_member_02.items()))) # 為了將新抽的人放最前面，所以要先倒序
                                    
                                    list_choose_keys = random.sample(list(dict_member_01.keys()), k = k_num) # 抽獎囉
                                    for key in list_choose_keys:
                                        dict_result[key] = dict_member_01.pop(key)   # 此時的dict_member_01已經更新
                                        dict_member_02[key] = dict_result[key]
                                    dict_member_02 = dict(reversed(list(dict_member_02.items()))) # 抽完之後再倒序
                                    dict_result = dict(reversed(list(dict_result.items())))

                                    extr_num += k_num # 已抽數量更新
                                    remain_num -= k_num # 剩餘數量更新
                                    
                                    if remain_num == 0:
                                        list_prize_01 = list_prize_01[1:]
                                        list_prize_02 = [current_prize] + list_prize_02
                                    else:
                                        pass

                                    # 更新內容
                                    str_update_member_01, str_update_member_02, str_update_prize_01, str_update_prize_02, str_update_result = '', '', '', '', ''
                                    
                                    if mode_prize == 0:
                                        current_prize_str = f'第 {count + 1} 個獎項'
                                        title = '{}(共 {} 名，實抽 {} 名，剩餘 {} 名)：'.format(current_prize_str, int(current_prize[-1]), k_num, remain_num)
                                        str_update_prize_01 = '\n'.join([i[0] for i in list_prize_01])
                                        str_update_prize_02 = '\n'.join([i[0] for i in list_prize_02]) 
                                    elif mode_prize == 1:
                                        current_prize_str = current_prize[0]
                                        title = '{}(共 {} 名，實抽 {} 名，剩餘 {} 名)：'.format(current_prize_str, int(current_prize[-1]), k_num, remain_num)
                                        str_update_prize_01 = '\n'.join([','.join(i) for i in list_prize_01])
                                        str_update_prize_02 = '\n'.join([','.join(i) for i in list_prize_02])
                                    elif mode_prize == 2:
                                        current_prize_str = ' 獎品內容：'.join(current_prize[0].split(','))
                                        title = '{}(共 {} 名，實抽 {} 名，剩餘 {} 名)：'.format(current_prize_str, int(current_prize[-1]), k_num, remain_num)
                                        str_update_prize_01 = '\n'.join([','.join(i) for i in list_prize_01])
                                        str_update_prize_02 = '\n'.join([','.join(i) for i in list_prize_02])
                                    else: # 應不會有mode_prize在{0, 1, 2}以外的可能
                                        pass

                                    if mode_member == 0:
                                        for key in dict_member_01:
                                            str_update_member_01 += f'{key}' + '\n'
                                        for key in dict_member_02:
                                            str_update_member_02 += f'{key}' + '\n'
                                        for key in dict_result:
                                            str_update_result += f'{key}' + '\n'
                                    elif mode_member in {1, 2, 3}:
                                        for key in dict_member_01:
                                            str_update_member_01 += f'{key},{dict_member_01[key]}' + '\n'
                                        for key in dict_member_02:
                                            str_update_member_02 += f'{key},{dict_member_02[key]}' + '\n'
                                        for key in dict_result:
                                            str_update_result += f'{key}, {dict_result[key]}' + '\n'
                                    else: # 應不會有mode_prize在{0, 1, 2, 3}以外的可能
                                        pass
                                    
                                    str_original_record = f'{title}\n{str_update_result}\n{str_original_record}'.strip()
                                    count = (count + 1) if remain_num == 0 else count
                                    ###################### 逐次開獎 End ######################

                                    window_input['extr_str'].update('已抽數量：{}'.format(extr_num))
                                    window_input['remain_str'].update('剩餘數量：{}'.format(remain_num))

                                    window_input['Mline_member_01'].Update(str_update_member_01.strip())
                                    window_input['Mline_member_02'].Update(str_update_member_02.strip())
                                    window_input['Mline_prize_01'].Update(str_update_prize_01.strip())
                                    window_input['Mline_prize_02'].Update(str_update_prize_02.strip())
                                    window_input['Mline_result'].Update(str_update_result.strip())
                                    window_input['Mline_allrecord'].Update(str_original_record.strip())
                                    print('開完')

                                    # 依照獎品產出資料要先確定剩餘為0，會用到mode_prize、current_prize_str、str_original_record
                                    # popup則需要mode_prize、current_prize_str、str_update_result

                                    # 儲存檔案
                                    if remain_num == 0:
                                        if input_values['output_mode'] == 'txt檔':
                                            now = datetime.datetime.now()
                                            filename = '逐次開獎結果_{}_{}.txt'.format(current_prize_str.split(' 獎品內容：')[0], now.strftime("%Y%m%d_%H%M%S"))
                                            with open(os.path.join(savefilepath, filename), 'a+', encoding = 'utf-8') as f:
                                                f.write('\n\n'.join([i for i in str_original_record.split('\n\n') if i.startswith(current_prize_str)]))
                                        elif input_values['output_mode'] == 'xlsx檔':
                                            # output_cn_member_list
                                            if mode_member == 0:
                                                output_cn_member_list = ['編號']
                                            elif mode_member == 1:
                                                output_cn_member_list = ['編號', '人名']
                                            elif mode_member == 2:
                                                output_cn_member_list = ['編號', '工號', '人名']
                                            elif mode_member == 3:
                                                output_cn_member_list = ['編號', '工號', '部門名稱', '人名']
                                            else:
                                                pass
                                            
                                            # output_cn_prize_list
                                            if mode_prize in [0, 1]:
                                                output_cn_prize_list = ['獎項']
                                            elif mode_prize == 2:
                                                output_cn_prize_list = ['獎項', '獎品內容']
                                            else:
                                                pass

                                            output_cn_list = output_cn_prize_list + output_cn_member_list

                                            output_list = []
                                            for prize_group_str in [i for i in str_original_record.split('\n\n') if i.startswith(current_prize_str)]:
                                                # output_prize_list
                                                output_prize_list = []
                                                if mode_prize in [0, 1]:
                                                    output_prize_list = [prize_group_str.split('\n')[0].split('(')[0]]
                                                elif mode_prize == 2:
                                                    output_prize_list = prize_group_str.split('\n')[0].split('(')[0].split(' 獎品內容：')
                                                else:
                                                    pass
                                                
                                                # output_member_list
                                                output_member_list = []
                                                for member in prize_group_str.split('\n')[1:]:
                                                    if mode_member in [0, 1, 2, 3]:
                                                        output_member_list = [i.strip() for i in member.split(',')]
                                                        output_list.append(output_prize_list + output_member_list)
                                            
                                            output_df = pd.DataFrame(output_list, columns = output_cn_list)
                                            now = datetime.datetime.now()
                                            filename = '逐次開獎結果_{}_{}.xlsx'.format(current_prize_str.split(' 獎品內容：')[0], now.strftime("%Y%m%d_%H%M%S"))
                                            StyleFrame(output_df).to_excel(os.path.join(savefilepath, filename), index = False, best_fit = output_cn_list).save()

                            window_input['Do'].update('準備', button_color = '#475841')
                        else:
                            pass   # print(window_input['Do'].get_text())

                    elif input_values['lottery_mode'] == '一次全開！！':
                        if window_input['Do'].get_text() == '抽獎':  # 確實是抽獎按鈕，一次全開絕對不需要準備
                            print('一次開獎')
                            str_original_record = input_values['Mline_allrecord'] # 先存取當下的開獎紀錄
                            for current_prize in list_prize_01:
                                ############## 一次全開 Start ############
                                dict_result = {}
                                dict_member_02 = dict(reversed(list(dict_member_02.items()))) # 為了將新抽的人放最前面，所以要先倒序
                                
                                list_choose_keys = random.sample(list(dict_member_01.keys()), k = int(current_prize[-1])) # 抽獎囉
                                for key in list_choose_keys:
                                    dict_result[key] = dict_member_01.pop(key)   # 此時的dict_member_01已經更新
                                    dict_member_02[key] = dict_result[key]
                                dict_member_02 = dict(reversed(list(dict_member_02.items()))) # 抽完之後再倒序
                                dict_result = dict(reversed(list(dict_result.items())))

                                list_prize_01 = list_prize_01[1:]
                                list_prize_02 = [current_prize] + list_prize_02

                                # 更新內容
                                str_update_member_01, str_update_member_02, str_update_prize_01, str_update_prize_02, str_update_result = '', '', '', '', ''
                                
                                if mode_prize == 0:
                                    current_prize_str = f'第 {count + 1} 個獎項'
                                    title = '{}(共 {} 名，實抽 {} 名，剩餘 {} 名)：'.format(current_prize_str, int(current_prize[-1]), int(current_prize[-1]), 0)
                                    str_update_prize_01 = '\n'.join([i[0] for i in list_prize_01])
                                    str_update_prize_02 = '\n'.join([i[0] for i in list_prize_02])
                                elif mode_prize == 1:
                                    current_prize_str = current_prize[0]
                                    title = '{}(共 {} 名，實抽 {} 名，剩餘 {} 名)：'.format(current_prize_str, int(current_prize[-1]), int(current_prize[-1]), 0)
                                    str_update_prize_01 = '\n'.join([','.join(i) for i in list_prize_01])
                                    str_update_prize_02 = '\n'.join([','.join(i) for i in list_prize_02])
                                elif mode_prize == 2:
                                    current_prize_str = ' 獎品內容：'.join(current_prize[0].split(','))
                                    title = '{}(共 {} 名，實抽 {} 名，剩餘 {} 名)：'.format(current_prize_str, int(current_prize[-1]), int(current_prize[-1]), 0)
                                    str_update_prize_01 = '\n'.join([','.join(i) for i in list_prize_01])
                                    str_update_prize_02 = '\n'.join([','.join(i) for i in list_prize_02])
                                else: # 應不會有mode_prize在{0, 1, 2}以外的可能
                                    pass

                                if mode_member == 0:
                                    for key in dict_member_01:
                                        str_update_member_01 += f'{key}' + '\n'
                                    for key in dict_member_02:
                                        str_update_member_02 += f'{key}' + '\n'
                                    for key in dict_result:
                                        str_update_result += f'{key}' + '\n'
                                elif mode_member in {1, 2, 3}:
                                    for key in dict_member_01:
                                        str_update_member_01 += f'{key},{dict_member_01[key]}' + '\n'
                                    for key in dict_member_02:
                                        str_update_member_02 += f'{key},{dict_member_02[key]}' + '\n'
                                    for key in dict_result:
                                        str_update_result += f'{key}, {dict_result[key]}' + '\n'
                                else: # 應不會有mode_prize在{0, 1, 2, 3}以外的可能
                                    pass
                                
                                str_original_record = f'{title}\n{str_update_result}\n{str_original_record}'.strip()
                                count += 1
                                ############## 一次全開 End ############

                        window_input['Mline_member_01'].Update(str_update_member_01.strip())
                        window_input['Mline_member_02'].Update(str_update_member_02.strip())
                        window_input['Mline_prize_01'].Update(str_update_prize_01.strip())
                        window_input['Mline_prize_02'].Update(str_update_prize_02.strip())
                        window_input['Mline_result'].Update(str_update_result.strip())
                        window_input['Mline_allrecord'].Update(str_original_record.strip())
                        # 一次全開的跳出視窗 可以直接用str_original_record解析 因為沒有單獎項分批抽的情況

                        # 儲存檔案
                        if input_values['output_mode'] == 'txt檔':
                            now = datetime.datetime.now()
                            filename = '一次抽完結果_{}.txt'.format(now.strftime("%Y%m%d_%H%M%S"))
                            with open(os.path.join(savefilepath, filename), 'a+', encoding = 'utf-8') as f:
                                f.write(str_original_record)
                        elif input_values['output_mode'] == 'xlsx檔':
                            # output_cn_member_list
                            if mode_member == 0:
                                output_cn_member_list = ['編號']
                            elif mode_member == 1:
                                output_cn_member_list = ['編號', '人名']
                            elif mode_member == 2:
                                output_cn_member_list = ['編號', '工號', '人名']
                            elif mode_member == 3:
                                output_cn_member_list = ['編號', '工號', '部門名稱', '人名']
                            else:
                                pass
                            
                            # output_cn_prize_list
                            if mode_prize in [0, 1]:
                                output_cn_prize_list = ['獎項']
                            elif mode_prize == 2:
                                output_cn_prize_list = ['獎項', '獎品內容']
                            else:
                                pass

                            output_cn_list = output_cn_prize_list + output_cn_member_list

                            output_list = []
                            for prize_group_str in str_original_record.split('\n\n'):
                                # output_prize_list
                                output_prize_list = []
                                if mode_prize in [0, 1]:
                                    output_prize_list = [prize_group_str.split('\n')[0].split('(')[0]]
                                elif mode_prize == 2:
                                    output_prize_list = prize_group_str.split('\n')[0].split('(')[0].split(' 獎品內容：')
                                else:
                                    pass
                                
                                # output_member_list
                                output_member_list = []
                                for member in prize_group_str.split('\n')[1:]:
                                    if mode_member in [0, 1, 2, 3]:
                                        output_member_list = [i.strip() for i in member.split(',')]
                                        output_list.append(output_prize_list + output_member_list)
                            
                            output_df = pd.DataFrame(output_list, columns = output_cn_list)
                            now = datetime.datetime.now()
                            filename = '一次抽完結果_{}.xlsx'.format(now.strftime("%Y%m%d_%H%M%S"))
                            StyleFrame(output_df).to_excel(os.path.join(savefilepath, filename), index = False, best_fit = output_cn_list).save()

                        print('開完')
                    else:
                        sg.PopupOK('怎麼有其他抽獎模式？', font = ('Microsoft YaHei', 10))
            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                error_message = '無法執行抽獎，請協助除錯。\n錯誤在第 {} 行，錯誤訊息：\n{}\n{}'.format(exc_tb.tb_lineno, exc_type, str(e))
                sg.PopupOK(error_message, font = ('Microsoft YaHei', 10))
            else:
                pass

window_input.close()
