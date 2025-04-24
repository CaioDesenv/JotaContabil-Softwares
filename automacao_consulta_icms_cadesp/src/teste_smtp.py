import smtplib
from email.message import EmailMessage

def testar_smtp():
    # Configurações do servidor SMTP
    config = {
        'host': 'cloud77.mailgrid.net.br',  # Exemplo de host SMTP
        'port': 465,                        # Porta para conexão via SMTP_SSL
        'usuario': 'jotacontabil@mail.jotacontabil.com.br',  # Usuário autenticado
        'senha': '3rwLJFIYzZ',              # Senha do SMTP
        'remetente': 'info@jotacontabil.com.br'  # E-mail remetente
    }

    # Cria a mensagem de e-mail
    mensagem = EmailMessage()
    mensagem['Subject'] = "Teste SMTP"
    mensagem['From'] = config.get("remetente")
    mensagem['To'] = config.get("remetente")  # Pode enviar para o próprio remetente para teste
    mensagem.set_content("Este é um e-mail de teste para verificar o funcionamento do servidor SMTP.")

    try:
        # Estabelece conexão utilizando SMTP_SSL
        with smtplib.SMTP_SSL(config.get("host"), config.get("port")) as servidor:
            servidor.login(config.get("usuario"), config.get("senha"))
            servidor.send_message(mensagem)
            print("E-mail de teste enviado com sucesso!")
    except Exception as erro:
        print(f"Erro ao enviar e-mail de teste: {erro}")

if __name__ == "__main__":
    testar_smtp()
