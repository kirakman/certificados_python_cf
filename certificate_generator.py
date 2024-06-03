from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import yagmail
from datetime import datetime
from config import pass_gmail

class EditCertificate:
    '''Classe que cria certificado dinamicamente e envia por email'''
    def __init__(self, user_email, user_password):
        self.user_email = user_email
        self.user_password = user_password
        self.read_data()
        self.send_email_with_certificate()

    def read_data(self):
        self.df = pd.read_excel('listaatual.xlsx')

    def send_email_with_certificate(self):
        for _, row in self.df.iterrows():
            name = row['Nome']  
            email = row['Email']  
            
            # Gerar um certificado com o nome e a data atual
            certificate_image = self.create_certificate(name)

            # Enviar email personalizado
            self.send_email_generic(name, email, certificate_image)

    def create_certificate(self, name):
        # Abrir a imagem de modelo
        imagem = Image.open('template.png')
        
        draw = ImageDraw.Draw(imagem)

        font_name = ImageFont.truetype('BRUSHSCI.TTF', 70)
        font_date = ImageFont.truetype('OPENSANS-VARIABLEFONT_WDTH,WGHT.ttf', 45)

        # Definindo a cor da fonte como preto (0, 0, 0)
        draw.text((1130, 545), name, font=font_name, fill=(0, 43, 89))

        # Adicionando a data atual no certificado
        current_date = datetime.now().strftime('%d/%m/%Y')
        draw.text((890, 847), f': {current_date}', font=font_date, fill=(0, 43, 89))

        certificate_image = f'{name}_certificate.png'
        imagem.save(certificate_image)
        return certificate_image

    def send_email_generic(self, name, email, certificate_image):
        # Inicializar o Yagmail SMTP
        usuario = yagmail.SMTP(user=self.user_email, password=self.user_password)
        
        assunto = 'Certificado de Participação - Segurança da Informação'
        conteudo = f'Olá {name},\n\nAqui está o seu certificado de participação. \n\nAtenciosamente,\nEquipe de Tecnologia'

        # Anexar a imagem do certificado
        usuario.send(
            to=email,
            subject=assunto,
            contents=conteudo,
            attachments=certificate_image
        )

        print(f'Email enviado para {email} com sucesso!')

# Substitua 'seu_email' e 'sua_senha' com o seu endereço de e-mail e senha do Gmail
start = EditCertificate('cyberfox.ads@gmail.com', 'exow knhu dmkt wmeh')
