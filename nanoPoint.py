import cv2
import pandas as pd
from datetime import datetime
import os
import tkinter as tk
from tkinter import simpledialog, messagebox

# Pasta para salvar fotos
if not os.path.exists("fotos"):
    os.makedirs("fotos")

# Planilha de registro
excel_file = "registro_ponto.xlsx"
if os.path.exists(excel_file):
    df = pd.read_excel(excel_file)
else:
    df = pd.DataFrame(columns=["ID", "Data", "Hora", "Tipo", "Foto"])

# Função principal
def registrar_ponto():
    global df  # precisamos atualizar o DataFrame global

    user_id = simpledialog.askstring("ID", "Digite seu ID:")
    if not user_id:
        return

    tipo = simpledialog.askstring("Tipo", "Digite o tipo: Entrada, Saída, Almoço, Retorno")
    if tipo not in ["Entrada", "Saída", "Almoço", "Retorno"]:
        messagebox.showerror("Erro", "Tipo inválido!")
        return

    # Capturar foto
    cap = cv2.VideoCapture(0)
    ret, frame = cap.read()
    cap.release()
    if not ret:
        messagebox.showerror("Erro", "Não foi possível acessar a câmera")
        return

    now = datetime.now()
    hora_str = now.strftime("%H:%M")
    data_str = now.strftime("%Y-%m-%d")
    foto_name = f"fotos/{user_id}_{data_str}_{hora_str.replace(':','-')}.jpg"
    cv2.imwrite(foto_name, frame)

    # Criar DataFrame temporário para o novo registro
    novo_registro = pd.DataFrame([{
        "ID": user_id,
        "Data": data_str,
        "Hora": hora_str,
        "Tipo": tipo,
        "Foto": foto_name
    }])

    # Adicionar ao DataFrame principal
    df = pd.concat([df, novo_registro], ignore_index=True)
    df.to_excel(excel_file, index=False)

    messagebox.showinfo("Sucesso", f"Ponto registrado: {tipo} às {hora_str}")

# Interface
root = tk.Tk()
root.withdraw()
while True:
    registrar_ponto()
