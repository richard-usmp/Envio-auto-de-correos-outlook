from base import send_email, add_body
from enumeraciones import ETipoEnvio

def send_emails(emisor,asunto,df, en, input_tuNombre):
    for receptor in df['EMAIL'].values:
        df_e = df[df['EMAIL']==receptor].T
        df_e.reset_index(level=0, inplace=True)
        df_e.columns = ['K','V']

        dic_kv = {}
        for i,r in df_e.iterrows():
            dic_kv[r['K']] = r['V']

        mail = send_email(emisor, receptor, asunto,dic_kv)

        if en == ETipoEnvio.AvanceCurso:

            porc = dic_kv['PORCENTAJE_AVANCE']
            if porc == 0:
                mail.CC=dic_kv['EMAIL_LIDER']
            #mail.BCC = 'danieldelgadof@bcp.com.pe'
            
            fname = 'Formatos\AvanceCursos.html'
            html_file = open(fname, 'r',encoding='utf-8')
            html = html_file.read()
            mail = add_body(html, mail, input_tuNombre, dic_kv)
            print("Correo enviado 1")

        elif en == ETipoEnvio.InicioMedicion: 
            
            #mail.CC=dic_kv['EMAIL_LIDER']
            fname = 'formatos\InicioMedicion.html'
            html_file = open(fname,'r', encoding='utf-8')
            html = html_file.read()
            mail = add_body(html, mail, input_tuNombre, dic_kv)
            print("Correo enviado 2")

        elif en == ETipoEnvio.InicioPDI:
            
            fname = 'Formatos\InicioPDI.html'
            html_file = open(fname,'r', encoding='utf-8')
            html = html_file.read()
            mail = add_body(html, mail, input_tuNombre, dic_kv)
            print("Correo enviado 3")

        mail.Send()
        #mail.Save()

    
        #mail = add_files(mail, dic=dic)      
