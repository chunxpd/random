from tkinter import Label, Button, Frame, Tk, Entry, messagebox, ttk
import random
from pandas import read_excel
from tabulate import tabulate
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# random sampling
data = read_excel("data.xlsx")
idx_data = [i for i in range(len(data.index))]
emp_id = "00000"
count = 0
msg = ""
num_result = 1

data = data.astype({'사번': 'str'})
data['num1'] = data['사번'].str.slice(start=0, stop=1)
data['num2'] = data['사번'].str.slice(start=1, stop=2)
data['num3'] = data['사번'].str.slice(start=2, stop=3)
data['num4'] = data['사번'].str.slice(start=3, stop=4)
data['num5'] = data['사번'].str.slice(start=4, stop=5)

raw_data = data


def get_entry(event=None):
    global num_result

    num = 1
    tmp = entry.get()
    if not tmp.isnumeric():
        messagebox.showerror("경고", "숫자를 입력해주세요!")
        num = 1
        entry.delete(0, len(tmp))
    else:
        num = int(tmp)
        if num <= 0:
            messagebox.showerror("경고", "양의 정수를 입력해주세요!!")
            num = 1
            entry.delete(0, len(tmp))
        elif num > data.shape[0]:
            messagebox.showerror("경고", "작은 숫자를 입력하세요!!")
            num = 1
            entry.delete(0, len(tmp))

    num_result = num
    label_entry.configure(text="뽑을 인원: " + str(num) + "명")
    entry.delete(0, len(tmp))
    entry.configure(state="disabled")


def func_filter_data(col_name, num_filter):
    global data
    data = data[data[col_name] == str(num_filter)]
    func_change_chart(data)


def func_reset_num1():
    global count
    global msg
    label_num1.configure(text="")
    label_num1.configure(text=emp_id[0], bg='lightblue')
    func_filter_data('num1', emp_id[0])
    count = count + 1
    if count == 7:
        result.configure(text=msg)


def func_reset_num2():
    global count
    global msg
    label_num2.configure(text="")
    label_num2.configure(text=emp_id[1], bg='lightblue')
    func_filter_data('num2', emp_id[1])
    count = count + 1
    if count == 7:
        result.configure(text=msg)


def func_reset_num3():
    global count
    global msg
    label_num3.configure(text="")
    label_num3.configure(text=emp_id[2], bg='lightblue')
    func_filter_data('num3', emp_id[2])
    count = count + 1
    if count == 7:
        result.configure(text=msg)


def func_reset_num4():
    global count
    global msg
    label_num4.configure(text="")
    label_num4.configure(text=emp_id[3], bg='lightblue')
    func_filter_data('num4', emp_id[3])
    count = count + 1
    if count == 7:
        result.configure(text=msg)


def func_reset_num5():
    global count
    global msg

    label_num5.configure(text="")
    label_num5.configure(text=emp_id[4], bg='lightblue')
    func_filter_data('num5', emp_id[4])
    count = count + 1
    if count == 7:
        result.configure(text=msg)

def func_reset_num6():
    global count
    global msg

    label_num5.configure(text="")
    label_num5.configure(text=emp_id[4], bg='lightblue')
    func_filter_data('num5', emp_id[4])
    count = count + 1
    if count == 7:
        result.configure(text=msg)

def func_reset_num7():
    global count
    global msg

    label_num5.configure(text="")
    label_num5.configure(text=emp_id[4], bg='lightblue')
    func_filter_data('num5', emp_id[4])
    count = count + 1
    if count == 7:
        result.configure(text=msg)

def func_pick():
    global emp_id
    global msg
    global num_result

    num = random.sample(idx_data, num_result)
    emp_id = str(data.iloc[num, 0].tolist()[0])

    print(num_result)

    tmp = data.iloc[num, 0:6]
    msg = "축하합니다!\n" + tabulate(tmp, showindex=False)

    if num_result == 1:
        result.configure(text="보기 버튼으로 숫자를 확인하세요!")
    else:
        result.configure(text=msg)

    button.configure(state="disabled")
    button_num1.configure(state="normal")
    button_num2.configure(state="normal")
    button_num3.configure(state="normal")
    button_num4.configure(state="normal")
    button_num5.configure(state="normal")

    entry.configure(state="disabled")


def func_change_chart(data):
    ax1.clear()
    df1 = data.groupby('호칭')['사번'].count().sort_values(ascending=True)
    df1.plot(kind='barh', legend=False, ax=ax1)
    ax1.set_title('호칭별 생존자 수')
    bar1.draw()


def init():
    global data
    global count
    global msg
    global emp_id
    global num_result

    data = raw_data
    count = 0
    msg = "뽑기 버튼을 눌러주세요!"
    emp_id = "00000"
    num_result = 1

    button.configure(state="normal")

    result.configure(text=msg)
    label_num1.configure(text="첫번째 자리", bg='white')
    label_num2.configure(text="두번째 자리", bg='white')
    label_num3.configure(text="세번째 자리", bg='white')
    label_num4.configure(text="네번째 자리", bg='white')
    label_num5.configure(text="다섯번째 자리", bg='white')

    button_num1.configure(state="disabled")
    button_num2.configure(state="disabled")
    button_num3.configure(state="disabled")
    button_num4.configure(state="disabled")
    button_num5.configure(state="disabled")

    entry.configure(state="normal")
    label_entry.configure(text="뽑을 인원: 1명")

    func_change_chart(raw_data)


# Variables
msg_title = "조합원 힐링지원 프로그램 뽑기 "
plt.rcParams['font.family'] = 'LG Smart_H2.0'

# tkinter
window = Tk()
window.title(msg_title)
window.config(padx=10, pady=10, bg="lightgrey")
window.geometry("550x550")
window.resizable(True, True)

# message frame
message_frame = Frame(window, width=800, height=80)
result = Label(message_frame, text="뽑기 버튼을 눌러주세요!", pady=10)
result.pack(pady=10)
message_frame.pack(pady=10, expand="no", fill="both", anchor="center")

# num result frame
num_result_frame = Frame(window, width=400, height=30)
label_entry = Label(num_result_frame, text="뽑을 인원: 1명", pady=10)
label_entry.grid(row=0)

entry = Entry(num_result_frame, width=10)
entry.bind("<Return>", get_entry)
entry.grid(row=1, column=0)

button_entry = Button(num_result_frame, text="입력", command=get_entry)
button_entry.grid(row=1, column=1)
num_result_frame.pack(pady=10, expand="no", anchor="center")

# button frame
bottom_frame = Frame(window, width=800, height=100)
button = Button(bottom_frame, text="숫자뽑기", bg="white", fg="black", \
                font=("LG Smart_H2.0", 12, "bold"), \
                command=func_pick)
button.grid(row=0, column=0)

button_init = Button(bottom_frame, text="초기화", bg="white", fg="red", \
                     font=("LG Smart_H2.0", 12, "bold"), \
                     command=init)
button_init.grid(row=0, column=1)

bottom_frame.pack(side="top", anchor="center")

# card frame
card_frame = Frame(window, width=800, height=100)
label_num1 = Label(card_frame, text="첫번째 자리", pady=3, \
                   font=("LG Smart_H2.0", 12, "bold"), bg="white", width=10, height=5)
label_num2 = Label(card_frame, text="두번째 자리", pady=3, \
                   font=("LG Smart_H2.0", 12, "bold"), bg="white", width=10, height=5)
label_num3 = Label(card_frame, text="세번째 자리", pady=3, \
                   font=("LG Smart_H2.0", 12, "bold"), bg="white", width=10, height=5)
label_num4 = Label(card_frame, text="네번째 자리", pady=3, \
                   font=("LG Smart_H2.0", 12, "bold"), bg="white", width=10, height=5)
label_num5 = Label(card_frame, text="다섯번째 자리", pady=3, \
                   font=("LG Smart_H2.0", 12, "bold"), bg="white", width=10, height=5)


button_num1 = Button(card_frame, text="보기", bg="white", fg="black", \
                     font=("LG Smart_H2.0", 12, "bold"), state="disabled", \
                     command=func_reset_num1)
button_num2 = Button(card_frame, text="보기", bg="white", fg="black", \
                     font=("LG Smart_H2.0", 12, "bold"), state="disabled", \
                     command=func_reset_num2)
button_num3 = Button(card_frame, text="보기", bg="white", fg="black", \
                     font=("LG Smart_H2.0", 12, "bold"), state="disabled", \
                     command=func_reset_num3)
button_num4 = Button(card_frame, text="보기", bg="white", fg="black", \
                     font=("LG Smart_H2.0", 12, "bold"), state="disabled", \
                     command=func_reset_num4)
button_num5 = Button(card_frame, text="보기", bg="white", fg="black", \
                     font=("LG Smart_H2.0", 12, "bold"), state="disabled", \
                     command=func_reset_num5)




label_num1.grid(row=0, column=0)
label_num2.grid(row=0, column=1)
label_num3.grid(row=0, column=2)
label_num4.grid(row=0, column=3)
label_num5.grid(row=0, column=4)


button_num1.grid(row=1, column=0)
button_num2.grid(row=1, column=1)
button_num3.grid(row=1, column=2)
button_num4.grid(row=1, column=3)
button_num5.grid(row=1, column=4)

card_frame.pack(pady=10, expand="no", fill="both", anchor="center")

import matplotlib

matplotlib.rcParams['font.family'] ='Malgun Gothic'

matplotlib.rcParams['axes.unicode_minus'] =False

# matplotlib
figure1 = plt.Figure(figsize=(30, 2), dpi=100)
ax1 = figure1.add_subplot(111)
bar1 = FigureCanvasTkAgg(figure1, window)
bar1.get_tk_widget().pack()
df1 = data.groupby('호칭')['사번'].count().sort_values(ascending=True)
df1.plot(kind='barh', legend=False, ax=ax1)
ax1.set_title('호칭별 생존자 수')

window.mainloop()