import tkinter as tk
from datetime import datetime
import telebot

# Configurações do bot do Telegram
API_KEY = '6049860098:AAGj7yTazEXn4yUFfcNJ44dh0pgJ4VYnOvA'
CHAT_ID = '-927904484'
NOMES_PERMITIDOS = ['Fernando', 'Isadora', 'Neto', 'Karina', 'Gabriel']
# Coloca em ordem alfabética
NOMES_PERMITIDOS.sort()

# Função para enviar mensagem para o Telegram
def enviar_telegram(mensagem):
    bot = telebot.TeleBot(token=API_KEY)
    bot.send_message(chat_id=CHAT_ID, text=mensagem)

def verificar_nome(nome):
    # Verifica se o nome realmente está na lista
    if nome in NOMES_PERMITIDOS:
        return True
    else:
        return False

# Função para registrar a entrada/saída do colaborador
def registrar_ponto():
    nome = var_nome.get()
    # Verifica se realmente possui um nome permitido
    if verificar_nome(nome):
        hora_atual = datetime.now().strftime("%H:%M:%S")
        data_atual = datetime.now().strftime(" - %d/%m/%Y")
        mensagem = f"{nome} registrou ponto às {hora_atual}"
        enviar_telegram(mensagem + data_atual)
        label_mensagem.configure(text=mensagem)
        # Reseta a opção para a inicial
        var_nome.set('----------')
    else:
        mensagem = "Selecione um nome válido!"
        label_mensagem.configure(text=mensagem)

# Função para atualizar o relógio
def atualizar_relogio():
    data_hora_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    label_data_hora.configure(text=data_hora_atual)
    root.after(1000, atualizar_relogio)

# Criação da janela principal
root = tk.Tk()
root.title("Controle de Ponto")

# Criação dos widgets
label_nome = tk.Label(root, text="Nome:")
var_nome = tk.StringVar(value="----------")
option_menu_nome = tk.OptionMenu(root, var_nome, *NOMES_PERMITIDOS)
button_registrar = tk.Button(root, text="Registrar", command=registrar_ponto)
label_mensagem = tk.Label(root, text="")
label_data_hora = tk.Label(root, font=('Arial', 18), pady=10)

# Posicionamento dos widgets na janela
label_nome.grid(row=0, column=0)
option_menu_nome.grid(row=0, column=1)
button_registrar.grid(row=1, column=0, columnspan=2)
label_mensagem.grid(row=2, column=0, columnspan=2)
label_data_hora.grid(row=3, column=0, columnspan=2)

# Atualização do relógio
atualizar_relogio()

# Inicialização da janela
root.mainloop()
