import tkinter as tk
from tkinter import messagebox
import os
import win32print
import win32ui
from win32con import *

# Inicializa contadores para as senhas
contador_padrao = 0
contador_preferencial = 0

def gerar_senha(tipo):
    global contador_padrao, contador_preferencial

    if tipo == "padrao":
        contador_padrao += 1
        senha = f"P{contador_padrao}"
    elif tipo == "preferencial":
        contador_preferencial += 1
        senha = f"A{contador_preferencial}"
    else:
        senha = ""

    return senha

def imprimir_senha(senha):
    try:
        # Definir a impressora padrão
        printer_name = win32print.GetDefaultPrinter()
        hprinter = win32print.OpenPrinter(printer_name)
        devmode = win32print.GetPrinter(hprinter, 2)["pDevMode"]

        # Configurar o dispositivo de impressão
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        hdc.StartDoc(senha)
        hdc.StartPage()

        # Definir a fonte com tamanho 30 pt (cerca de 60 unidades no sistema)
        font = win32ui.CreateFont({
            "name": "Arial", 
            "height": 60,  # Tamanho 30 pt
            "weight": 700,  # Negrito
        })
        hdc.SelectObject(font)

        # Escrever o texto na página
        hdc.TextOut(100, 100, f"Senha: {senha}")

        # Finalizar a página
        hdc.EndPage()
        hdc.EndDoc()
        hdc.DeleteDC()

        messagebox.showinfo("Impressão", f"Senha {senha} enviada para a impressão!")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao imprimir: {e}")

def gerar_e_imprimir(tipo):
    senha = gerar_senha(tipo)
    senha_label.config(text=senha)
    imprimir_senha(senha)

# Criação da janela principal
janela = tk.Tk()
janela.title("Gerador de Senhas")
janela.geometry("300x200")

# Layout da interface
titulo_label = tk.Label(janela, text="Clique no botão para gerar a senha:")
titulo_label.pack(pady=10)

btn_padrao = tk.Button(janela, text="Senha Padrão", command=lambda: gerar_e_imprimir("padrao"))
btn_padrao.pack(pady=5)

btn_preferencial = tk.Button(janela, text="Senha Preferencial", command=lambda: gerar_e_imprimir("preferencial"))
btn_preferencial.pack(pady=5)

senha_texto_label = tk.Label(janela, text="Senha Gerada:")
senha_texto_label.pack(pady=5)

senha_label = tk.Label(janela, text="", font=("Arial", 14))
senha_label.pack(pady=5)

btn_sair = tk.Button(janela, text="Sair", command=janela.quit)
btn_sair.pack(pady=10)

# Loop principal
janela.mainloop()
