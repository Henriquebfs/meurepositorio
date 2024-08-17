
import pandas as pd
import smtplib
import email.message

#importando planilha em excel com os dados
planilha = pd.read_excel("Rastreios Elmeco.xlsx")

#planilha.info()  -  Para verificar qtd e os nomes das colunas

print("--- ENVIO DE E-MAILS PERSONALIZADOS DE ACORDO COM A PLANILHA: Rastreios Elmeco.xlsx ---\n") #Titulo do aplicativo

remetente = input("Digite o e-mail que será usado como remetente: ")
senha = input("Digite a senha de aplicativo para autorização do Gmail: ")
confirma = input("Tem certeza que deseja enviar os e-mails para todos os destinatários conforme a planilha?\n Digite S ou N e pressione ENTER: ")

print("\n Aguarde...\n")

if((confirma == "N") or (confirma == "n")):
    print("Operação Cancelada.")

elif((confirma == "S") or (confirma =="s")):
    for i, mail in enumerate(planilha['E-MAIL DESTINATARIO']):
        cliente = planilha.loc[i, 'CLIENTE DESTINATARIO']
        pedido = planilha.loc[i, 'NUMERO PEDIDO']
        data = planilha.loc[i, 'DATA']
        rastreio = planilha.loc[i, 'AWB']

        corpo_email = f"""
        <p>Prezado(a) {cliente},</p>
        <p>Seu pedido foi enviado via GOLLOG conforme dados abaixo:</p>
        <pre><b> PEDIDO           DATA               RASTREIO</b></pre>
        <pre> {pedido}         {data}           {rastreio}</pre> 


        <p>Segue o link para acompanhar o rastreio do seu pedido:</p>
        https://servicos.gollog.com.br/app/main/tracking
        <br />
        <p>Em caso de dúvidas, estamos à disposição.</p>
        <div>
        <img src = "https://docs.google.com/uc?id=1s7kMhLSF8qzHvZT7ulz-cdOeZzOvKRW7">
        </div>
        """

        msg = email.message.Message()
        msg['Subject'] = "ELMECO - Código de rastreio"
        msg['From'] = remetente # E-mail que será usado como remetente
        msg['To'] = mail
        password = senha #especificar uma senha de app em https://security.google.com/settings/security/apppasswords
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email )

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('Email enviado')

else:
    print('Operação Cancelada')
     
input("\nPronto. Programa Finalizado. Pressione ENTER para sair.")

