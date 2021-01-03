from tkinter import *
from tkinter import filedialog
from app import check_if_on_reminder_list, check_same_price_within_6_days, save_results, load_excel

window = Tk()

window.title('ISIN nalyser 2.0')
window.geometry('350x200')

lbl = Label(window, text='ISIN analyse')
lbl.grid(column=0, row=0)


def load_file():
    lbl.configure(text='Yolo')
    file = filedialog.askopenfile()
    if file:
        print(file.name)
        df_instruments, df_internal_data, df_reminder_list, df_comments, df_cbb = load_excel(file.name)

        check_if_on_reminder_list(df_instruments, df_reminder_list)
        check_same_price_within_6_days(df_instruments, df_internal_data, '22-05-2020')

        btn_save = Button(window, text='Save results',
                          command=lambda: save_excel(df_instruments, df_internal_data, df_reminder_list, df_comments,
                                                     df_cbb,
                                                     file.name))
        btn_save.grid(column=3, row=0)


def save_excel(df_instruments, df_internal_data, df_reminder_list, df_comments, df_cbb, path):
    save_results(df_instruments, df_internal_data, df_reminder_list, df_comments, df_cbb, path)


btn = Button(window, text='Load File', command=load_file)
btn.grid(column=2, row=0)

window.mainloop()
