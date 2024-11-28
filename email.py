import win32com.client as win32

try:
    # Criar a integração com o Outlook
    outlook = win32.Dispatch('outlook.application')
    
    # Criar um e-mail
    email = outlook.CreateItem(0)
    
    # Configurar as informações do e-mail
    email.To = "victorfp335@gmail.com;"
    email.Subject = "E-mail automático do Python"
    email.HTMLBody = '''
    <p>Caro usuário,</p>
    <p>Precisamos que você altere a sua senha seguindo a sua permissão da recuperação de sua senha.</p>
    <p>Atenciosamente,<br>A equipe do Larlocker</p>
    '''
    
    # Alterar o remetente (e-mail de envio)
    email.SentOnBehalfOfName = "larlocker.pi@gmail.com"  # Substitua pelo e-mail desejado
    
    # Enviar o e-mail
    email.Send()
    print("Email enviado com sucesso.")
    
except Exception as e:
    print(f"Erro ao enviar o e-mail: {e}")
