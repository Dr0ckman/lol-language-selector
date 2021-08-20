import os
import re
import tkinter as tk
import winshell
import win32com.client
from os import chmod
from pathlib import Path
from tkinter import filedialog
from lang_dict import lang_dict


root = tk.Tk()
root.title("LoL Language Selector")
file_path = ""
chosen_lang = ""

# frames
config = tk.LabelFrame(root, text="Configuración", padx=5, pady=5)
config.pack(padx=5, pady=5, expand=True, fill="both")
buttons = tk.LabelFrame(root, text="Aplicar cambios", padx=5, pady=5)
buttons.pack(padx=5, pady=5, expand=True, fill="both")

# labels
label1 = tk.Label(config, text="Ubicación de LeagueClientSettings.yaml:", width=53)
label1.grid(column=0, row=0)

label2 = tk.Label(config, text="Idioma: ")
label2.grid(column=0, row=2)

label3 = tk.Button(config, text="", relief='sunken', state='disabled', width=53)
label3.grid(column=1, row=0)

statusbar = tk.Label(root, text="Listo", bd=1, relief=tk.SUNKEN, anchor=tk.W, bg="#b2c1d3")
statusbar.pack(side=tk.BOTTOM, fill=tk.X)


# buttons
def get_file():
    global file_path
    file_path = filedialog.askopenfilename(initialdir="/Riot Games/League of Legends/Config/",
                                           filetypes=(("yaml files", "*.yaml"), ("all files", "*.*"),))
    label3.configure(text=file_path)


select_dir = tk.Button(config, text="Buscar archivo", command=get_file, width=53)
select_dir.grid(column=1, row=1)

lang_var = tk.StringVar()
lang = tk.OptionMenu(config, lang_var, "Alemán", "Checo", "Español", "Francés", "Griego",
                     "Húngaro", "Inglés", "Italiano", "Japonés", "Polaco", "Portugués",
                     "Rumano", "Ruso", "Turco")
lang.grid(column=1, row=2, sticky="WE")

perm_var = tk.BooleanVar()  # Variable de change_permissions
shortcut_var = tk.BooleanVar()  # Variable de create_shortcut

change_permissions = tk.Checkbutton(config, text="Cambiar permisos a sólo-lectura ",
                                    variable=perm_var)
change_permissions.grid(column=1, row=3)


def create_shortcut():
    try:
        exe_path = Path(file_path)
        exe_path_parent = exe_path.parents[1]

        desktop = winshell.desktop()
        desktop_path = os.path.join(desktop, "League of Legends.lnk")
        target = os.path.join(exe_path_parent, "LeagueClient.exe")
        wDir = exe_path_parent

        shell = win32com.client.Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(str(desktop_path))
        shortcut.Targetpath = str(target)
        shortcut.WorkingDirectory = str(wDir)
        shortcut.Arguments = "-locale=" + str(chosen_lang)
        statusbar.configure(text="Creando acceso directo...")
        shortcut.save()
        statusbar.configure(text="Acceso directo creado correctamente.")

        if perm_var is True:
            statusbar.configure(text="Cambiando permisos de acceso directo")
            chmod(str(desktop_path), 100)
            statusbar.configure(text="Acceso directo configurado como sólo-lectura")
        else:
            pass
    except IndexError:
        statusbar.configure(text="Seleccione la ubicación del archivo LeagueClientSettings.yaml")
    # except:
    #     statusbar.configure(text="No se pudo guardar el acceso directo. Verifique que no exista un acceso directo.")

def change_language():
    # global chosen_lang
    try:
        # for language in lang_dict:
        #     if language == lang_var.get():
        #         chosen_lang = lang_dict[language]
        with open(file_path, "r+") as file:
            f_contents = file.read()
            regex = re.sub('[a-z][a-z]_[A-Z][A-Z]', chosen_lang, f_contents)
            file.seek(0)
            statusbar.configure(text="Configurando LeagueClientSettings.yaml...")
            file.write(regex)
            file.truncate()
            statusbar.configure(text="Idioma cambiado correctamente.")
        if perm_var.get() is True:
            statusbar.configure(text="Cambiando permisos...")
            chmod(file_path, 100)
            statusbar.configure(text="LeagueClientSettings.yaml configurado como sólo-lectura")
        else:
            pass
    except PermissionError:
        statusbar.configure(text="Archivo de sólo-lectura. Cambie los permisos manualmente")
    except FileNotFoundError:
        statusbar.configure(text="Seleccione la ubicación del archivo LeagueClientSettings.yaml y un idioma")


def accept():
    global chosen_lang
    for language in lang_dict:
        if language == lang_var.get():
            chosen_lang = lang_dict[language]
    create_shortcut()
    change_language()


# c_shortcut_button = tk.Button(buttons, text="Crear acceso directo", command=create_shortcut, width=53)
# c_shortcut_button.grid(column=0, row=5)

accept_button = tk.Button(buttons, text="Cambiar idioma", command=accept, width=108)
accept_button.grid(column=1, row=5, columnspan=1)

root.mainloop()
