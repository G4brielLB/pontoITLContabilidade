import tkinter as tk
from datetime import datetime
import telebot

# Configurações do bot do Telegram
API_KEY = '6049860098:AAGj7yTazEXn4yUFfcNJ44dh0pgJ4VYnOvA'
CHAT_ID = '-927904484'

# Função para enviar mensagem para o Telegram
def enviar_telegram(mensagem):
    bot = telebot.TeleBot(token=API_KEY)
    bot.send_message(chat_id=CHAT_ID, text=mensagem)

def verificar_nome(nome):
    """
    Verifica se o nome está escrito apenas com letras (sem números e outros caracteres).
    A função .isalpha() do Python verifica tudo, inclusive os espaços entre os nomes, o que não queremos.
    Logo, os espaços entre os nomes serão trocados por um caracter alfabético, no caso, 'a'.
    Aí sim a função .isalpha() funcionará bem, apenas conferindo se todos os caracteres são alfabéticos.
    Se tiver tudo certo, retorna True, se tiver algum caracter inválido, retorna False
    """
    verificador = nome.find(" ")
    # Se não tiver nenhum espaço entre os nomes
    if verificador == -1:
        if nome.isalpha():
            return True
        else:
            return False
    # Se tiver espaço entre os nomes
    else:
        nome = nome.replace(" ", "a")
        if nome.isalpha():
            return True
        else:
            return False

# Função para registrar a entrada/saída do colaborador
def registrar_ponto():
    nome = entry_nome.get()
    # Remove os espaços indevidos antes e depois do nome
    nome = nome.strip()
    # Verifica se realmente possui um nome escrito e possui apenas caracteres alfabéticos
    if verificar_nome(nome):
        hora_atual = datetime.now().strftime("%H:%M:%S")
        data_atual = datetime.now().strftime(" - %d/%m/%Y")
        mensagem = f"{nome} registrou ponto às {hora_atual}"
        enviar_telegram(mensagem + data_atual)
        label_mensagem.configure(text=mensagem)
        # Remove o nome anteriormente escrito na caixinha de texto
        entry_nome.delete(0, len(nome))
    else:
        mensagem = "Digite um nome válido!"
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
entry_nome = tk.Entry(root)
button_registrar = tk.Button(root, text="Registrar", command=registrar_ponto)
label_mensagem = tk.Label(root, text="")
label_data_hora = tk.Label(root, font=('Arial', 18), pady=10)


# Posicionamento dos widgets na janela
label_nome.grid(row=0, column=0)
entry_nome.grid(row=0, column=1)
button_registrar.grid(row=1, column=0, columnspan=2)
label_mensagem.grid(row=2, column=0, columnspan=2)
label_data_hora.grid(row=3, column=0, columnspan=2)

# Atualização do relógio
atualizar_relogio()

# Inicialização da janela
root.mainloop()
