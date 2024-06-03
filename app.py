from flask import Flask, request, jsonify
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime
import yagmail
from config import pass_gmail

app = Flask(__name__)

class EditCertificate:
    def __init__(self, user_email, user_password):
        self.user_email = user_email
        self.user_password = user_password

    def read_data(self):
        self.df = pd.read_excel('listaatual.xlsx')

    def send_email_with_certificate(self):
        for _, row in self.df.iterrows():
            name = row['Nome']
            email = row['Email']
            
            certificate_image = self.create_certificate(name)
            self.send_email_generic(name, email, certificate_image)

    def create_certificate(self, name):
        imagem = Image.open('template.png')
        
        draw = ImageDraw.Draw(imagem)

        font_name = ImageFont.truetype('BRUSHSCI.TTF', 70)
        font_date = ImageFont.truetype('OPENSANS-VARIABLEFONT_WDTH,WGHT.ttf', 45)

        draw.text((1130, 545), name, font=font_name, fill=(0, 43, 89))
        current_date = datetime.now().strftime('%d/%m/%Y')
        draw.text((890, 847), f': {current_date}', font=font_date, fill=(0, 43, 89))

        certificate_image = f'{name}_certificate.png'
        imagem.save(certificate_image)
        return certificate_image

    def send_email_generic(self, name, email, certificate_image):
        usuario = yagmail.SMTP(user=self.user_email, password=self.user_password)
        
        assunto = 'Certificado de Participação - Segurança da Informação'
        conteudo = f'Olá {name},\n\nAqui está o seu certificado de participação. \n\nAtenciosamente,\nEquipe de Tecnologia'

        usuario.send(
            to=email,
            subject=assunto,
            contents=conteudo,
            attachments=certificate_image
        )

        print(f'Email enviado para {email} com sucesso!')

@app.route('/send-certificates', methods=['POST'])
def send_certificates():
    data = request.json
    email = data['email']
    password = data['password']
    certificate_service = EditCertificate(email, password)
    certificate_service.send_email_with_certificate()
    return jsonify({'message': 'Certificados enviados com sucesso!'})

if __name__ == '__main__':
    app.run(debug=True)
