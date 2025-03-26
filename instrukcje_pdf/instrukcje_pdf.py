import os
import sys
import tkinter as tk
from functools import partial
import sqlanydb
import keyring
import tkinter.simpledialog
from tkinter import messagebox
import logging
from pathlib import Path
from PIL import ImageTk, Image
import os.path as path
import time
os.environ['PYGAME_HIDE_SUPPORT_PROMPT'] = "hide"
from pygame import mixer

results = []
window = tk.Tk()
#show = 0


def is_file_older_than_x_days(file, days=1):
    file_time = path.getmtime(file)
    # Check against 24 hours
    return (time.time() - file_time) / 3600 > 24*days


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def get_pass():
    try:
        # tk.Tk().withdraw()
        window.withdraw()
        p = tkinter.simpledialog.askstring("Password", "Wprowadź hasło do bazy danych:", show='*')
        keyring.set_password("db", "127cfe29c39c", p)
        window.deiconify()
    except Exception as _:
        messagebox.showerror('Error', f'{_} in >get_pass< module')


def open_file(result):
    try:
        #print(f'OTU{result}')
        os.startfile(f'{result[3]}\\{result[2]}')
    except Exception as _:
        messagebox.showerror('Error', f'{_} in >open_file< module')


def search(_):
    results.clear()
    alert = 0
    # monodbc
    try:
        con = sqlanydb.connect(UID=keyring.get_password("db", "127cfe29c39c"), PWD=keyring.get_password("db", "127cfe29c39c"), Host='192.168.0.3', Server='klpl-monwin',
                           DBF='e:\\monwin\\db\\001\\monitor.db', DBN='FTG_001')
        query = (
            f"SELECT art_artnr, art_ben, fil_namn, svag_path FROM monitor.FIL LEFT outer join monitor.FILTEXT ON "
            f"FIL.fil_nr = FILTEXT.fil_nr left outer join monitor.SOKVAG ON SOKVAG.svag_nr = FIL.svag_nr left outer join "
            f"monitor.ARTIKEL on art_g_text = FILTEXT.tn_nr WHERE art_artnr='{_.get()}'")
        # print(entry.widget.get())
        # 848003
        _.delete(0, tk.END)
        cursor = con.cursor()
        cursor.execute(query)

        # print(cursor.description)
        rows = cursor.fetchall()

        cursor.close()
        con.close()

        for row in rows:
            # print(row)
            results.append(row)
        # choose_frame = tk.Frame(master=window)
        choose_frame.grid(row=2, column=0)
        label = tk.Label(master=choose_frame, text="Znalezione:")
        label.pack()

        for result in results:
            path = f'{result[3]}\\{result[2]}'
            selection = tk.Frame(master=choose_frame)
            selection.pack()
            action = partial(open_file, result)
            label = tk.Button(master=choose_frame, text=result[2], command=action, bg="blue")
            try:
                is_older = is_file_older_than_x_days(path, 14)
                if is_older:
                    label["bg"] = "green"
                if not is_older:
                    label["bg"] = "red"
                    alert = 1
                if alert:
                    label1.grid(row=3, column=0)
                    sound.play()
            except Exception as _:
                pass
                #messagebox.showerror('Error', f'{_} in >search: labels colors< module')
            label.pack()

        btn_searchFile["state"] = "disable"
        entry["state"] = "disable"
    except Exception as _:
        messagebox.showerror('Error', f'{_} in >search< module')


def clear():
    try:
        btn_searchFile["state"] = "normal"
        for widget in choose_frame.winfo_children():
            widget.destroy()
        choose_frame.grid_forget()
        results.clear()
        entry["state"] = "normal"
        sound.stop()
        label1.grid_forget()
    except Exception as _:
        messagebox.showerror('Error', f'{_} in >clear< module')


def clear_pass():
    try:
        keyring.delete_password("db", "127cfe29c39c")
        get_pass()
        pass
    except Exception as _:
        messagebox.showerror('Error', f'{_} in >clear_pass< module')


mixer.init()
sound = mixer.Sound(resource_path("siren.mp3"))


try:
    logfile = f'{(Path.home())}\\PDFlog.txt'
    if os.path.exists(logfile):
        pass
    else:
        open(logfile, "w")
    logging.basicConfig(filename=logfile,
                        filemode='a',
                        format='%(asctime)s,%(msecs)d %(levelname)s %(message)s',
                        datefmt="%Y-%m-%d %H:%M:%S",
                        level=logging.NOTSET
                        )

    logging.info("")
    logging.info("----------------------------")
    logging.info("********** START **********")

    if not keyring.get_password("db", "127cfe29c39c"):
        get_pass()

    window.columnconfigure(2, minsize=100, weight=1)
    window.rowconfigure(3, minsize=200, weight=1)

    choose_frame = tk.Frame(master=window)
    choose_frame.grid(row=1, column=0)

    entry = tk.Entry()
    entry.insert(0, "848003")
    entry.grid(row=0, column=0)
    entry.bind("<Return>", lambda event: search(entry))

    ac = partial(search, entry)
    btn_searchFile = tk.Button(master=window, text="Wyszukaj", command=ac)
    btn_searchFile.grid(row=0, column=2)

    btn_clear = tk.Button(master=window, text="Wyczyść", command=clear)
    btn_clear.grid(row=2, column=2)

    btn_clear_pass = tk.Button(master=window, text="Usuń hasło", command=clear_pass)
    btn_clear_pass.grid(row=3, column=2)

    image1 = Image.open(resource_path("g.jpg"))
    image1 = image1.resize((100, 100), Image.NEAREST)
    test = ImageTk.PhotoImage(image1)
    label1 = tkinter.Label(image=test)
    label1.grid(row=3, column=0)

    window.mainloop()
    # keyring.delete_password("db", "127cfe29c39c")
except Exception as _:
    messagebox.showerror('Error', f'{_} in >main< module')
