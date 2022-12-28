import PySimpleGUI as sg

def nameplate_layout(text):
    return sg.Text(
        text, 
        size = (30, 1),
        justification = 'center', 
        font = ('Microsoft YaHei', 18), 
        relief = sg.RELIEF_RIDGE
    )

def winners_list_layout(members_str, align_num):
	x, y = [nameplate_layout(', '.join([j.strip() for j in mb.split(',')])) for mb in members_str.strip().split('\n')], []
	for i in list(range(len(x)//align_num + 1)):
		y.append(x[(i * align_num):((i + 1) * align_num)])
	return y

def prize_frame_layout(mline_str, align_number, prize_title):
    # mline_str可放多行字串
    return sg.Frame(
        layout = [
            [sg.Column(
                winners_list_layout(members_str = mline_str, align_num = align_number), 
                element_justification = 'center'
            )],
        ],
        title = prize_title,
        relief = sg.RELIEF_GROOVE,
        font = ('Microsoft YaHei', 15),
        title_color = '#C00000',
        element_justification = 'center',
    )