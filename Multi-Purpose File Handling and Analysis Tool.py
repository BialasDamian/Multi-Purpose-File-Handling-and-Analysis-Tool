import os
import subprocess
import tkinter as tk
import tkinter.messagebox as msg
from tkinter import filedialog, ttk

import pandas as pd
from openpyxl import Workbook
from PIL import Image, ImageTk


def pokaz_strone(strona):
    # Ukryj wszystkie strony
    for page in strony:
        page.grid_forget()

    # Pokaż wybraną stronę
    strona.grid()


def miejsce_zapisu():
    global folder_path
    folder_path = filedialog.askdirectory()


def zapisz_plik():
    global nazwa, folder_path
    nazwa = ent_nazwa_dokumentu.get()
    if not nazwa:
        msg.showwarning("BŁĄD", "!!! MUSISZ PODAĆ NAZWĘ PLIKU !!!")
    elif not folder_path:
        msg.showwarning("BŁĄD", "!!! MUSISZ WYBRAĆ MIEJSCE ZAPISU PLIKU !!!")
    else:
        workbook = Workbook()
        workbook.save(f"{folder_path}/{nazwa}.xlsx")
        msg.showinfo("Sukces", "Plik został zapisany.")


def uruchom_plik():
    global sciezka_pliku
    nazwa1 = ent_nazwa_pliku.get()
    if nazwa1:
        znajdz_plik(nazwa1 + ".xlsx")
    else:
        msg.showwarning("BŁĄD", "!!! NIEPODANA NAZWA PLIKU !!!")

    try:
        subprocess.Popen([sciezka_pliku], shell=True)
        print("Plik został uruchomiony.")
        ent_nazwa_dokumentu.delete(0, tk.END)
    except Exception as e:
        print("Wystąpił błąd podczas uruchamiania pliku:", e)


def znajdz_plik(nazwa_pliku):
    global sciezka_pliku
    # Przeszukiwanie wszystkich dysków w systemie
    for root, dirs, files in os.walk('C:\\'):
        # Przeszukiwanie plików w bieżącym folderze
        for file in files:
            # Sprawdzenie, czy nazwa pliku pasuje do szukanej nazwy
            if file == nazwa_pliku:
                # Znaleziono plik, możesz zrobić coś z nim
                sciezka_pliku = os.path.join(root, file)
                print("Znaleziono plik:", sciezka_pliku)


def nazwa_pliku():
    global nazwa
    nazwa = ent_nazwa_dokumentu.get()
    if not nazwa:
        msg.showwarning("BŁĄD", "!!! MUSISZ PODAĆ NAZWĘ PLIKU !!!")


def znajdz_do_zamiany(nazwa):
    # Przeszukiwanie wszystkich dysków w systemie
    for root, dirs, files in os.walk('C:\\Users\Damian\Documents'):
        # Przeszukiwanie plików w bieżącym folderze
        for file in files:
                        # Sprawdzenie, czy nazwa pliku pasuje do szukanej nazwy i jest plikiem Excela
            if file == nazwa and file.endswith('.xlsx'):
                # Znaleziono plik, możesz zrobić coś z nim
                sciezka_pliku = os.path.join(root, file)
                print("Znaleziono plik:", sciezka_pliku)


# Tworzenie okna głównego
okno = tk.Tk()
okno.title("Program")
okno.geometry('768x512')
obrazek = tk.PhotoImage(file=os.path.abspath("D:\Projekty_Github\Obrazek1.png"))

# Tworzenie stron
strona_glowna = tk.Frame(okno)
strona_zapisu = tk.Frame(okno)
strona_uruchomienia = tk.Frame(okno)
strona_analizy_danych = tk.Frame(okno)
# Inicjalizacja zmiennych
strony = [strona_glowna, strona_zapisu, strona_uruchomienia,strona_analizy_danych]
nazwa = ""
folder_path = ""
sciezka_pliku = ""

# Strona główna

lbl_obrazek = tk.Label(strona_glowna,image=obrazek)
lbl_obrazek.grid(row=0, column=0, rowspan=3, columnspan=3, sticky="nsew")
lbl_witaj = tk.Label(strona_glowna, text="Menu główne",font =("Arial",30,"bold"))
lbl_witaj.place(relx=0.5, rely=0.2,anchor="center")

btn_strona_zapisu = tk.Button(strona_glowna, text="Baza danych",font=("Arial",14,"bold"), command=lambda: pokaz_strone(strona_zapisu))
btn_strona_zapisu.place(relx=0.5, rely=0.5,anchor="center")

btn_strona_uruchomienia = tk.Button(strona_glowna, text="Uruchom plik",font=("Arial",14,"bold"),
                                    command=lambda: pokaz_strone(strona_uruchomienia))
btn_strona_uruchomienia.place(relx=0.5, rely=0.6,anchor="center")

btn_strona_analizy_danych = tk.Button(strona_glowna, text="Analiza danych",font=("Arial",14,"bold"),
                                    command=lambda: pokaz_strone(strona_analizy_danych))
btn_strona_analizy_danych.place(relx=0.5, rely=0.7,anchor="center")

# Strona zapisu
lbl_obrazek_strona_zapisu = tk.Label(strona_zapisu,image=obrazek)
lbl_obrazek_strona_zapisu.grid(row=0, column=0, rowspan=3, columnspan=3, sticky="nsew")

lbl_nazwa_dokumentu = tk.Label(strona_zapisu, text="Podaj nazwę dokumentu:",font=("Arial",12,"bold"))
lbl_nazwa_dokumentu.place(relx=0.5, rely=0.3,anchor="center")

ent_nazwa_dokumentu = tk.Entry(strona_zapisu,font=("Arial",12))
ent_nazwa_dokumentu.place(relx=0.5, rely=0.35,anchor="center",)

btn_wybierz_folder = tk.Button(strona_zapisu, text="Wybierz folder",font=("Arial",12,"bold"), command=miejsce_zapisu)
btn_wybierz_folder.place(relx=0.5, rely=0.4,anchor="center")

btn_zapisz_plik = tk.Button(strona_zapisu, text="Zapisz plik",font=("Arial",12,"bold"), command=zapisz_plik)
btn_zapisz_plik.place(relx=0.5, rely=0.45,anchor="center")

btn_wroc_do_glownej1 = tk.Button(strona_zapisu, text="Wróć do strony głównej",font=("Arial",12,"bold"), command=lambda: pokaz_strone(strona_glowna))
btn_wroc_do_glownej1.place(relx=0.5, rely=0.5,anchor="center")

# Strona uruchomienia
lbl_obrazek_strona_uruchomienia = tk.Label(strona_uruchomienia,image=obrazek)
lbl_obrazek_strona_uruchomienia.grid(row=0, column=0, rowspan=3, columnspan=3, sticky="nsew")

lbl_nazwa_pliku = tk.Label(strona_uruchomienia, text="Podaj nazwę pliku do uruchomienia:",font=("Arial",12,"bold"))
lbl_nazwa_pliku.place(relx=0.5, rely=0.3,anchor="center")

ent_nazwa_pliku = tk.Entry(strona_uruchomienia,font=("Arial",12))
ent_nazwa_pliku.place(relx=0.5, rely=0.35,anchor="center")

btn_uruchom_plik = tk.Button(strona_uruchomienia, text="Uruchom plik",font=("Arial",12,"bold"), command=uruchom_plik)
btn_uruchom_plik.place(relx=0.5, rely=0.4,anchor="center")

btn_wroc_do_glownej2 = tk.Button(strona_uruchomienia, text="Wróć do strony głównej",font=("Arial",12,"bold"),
                                 command=lambda: pokaz_strone(strona_glowna))
btn_wroc_do_glownej2.place(relx=0.5, rely=0.45,anchor="center")

#Strona analizy danych
lbl_obrazek_strona_analizy_danych = tk.Label(strona_analizy_danych,image=obrazek)
lbl_obrazek_strona_analizy_danych.grid(row=0, column=0, rowspan=3, columnspan=3, sticky="nsew")

lbl_jaki_dokument = tk.Label(strona_analizy_danych,text="Jaki dokument chcesz poddać analizie?",font=("arial",12,"bold"))
lbl_jaki_dokument.place(relx=0.5, rely=0.3,anchor="center")


# Wyświetlanie strony głównej
pokaz_strone(strona_glowna)

# Uruchomienie pętli głównej programu
okno.mainloop()
