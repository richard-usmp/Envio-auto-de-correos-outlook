from base import send_email, add_body, add_body_with_image, add_files
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

        if en == ETipoEnvio.AvanceCurso_0_porc_CC_lider:

            porc = dic_kv['PORCENTAJE_AVANCE']
            if porc == 0:
                mail.CC=dic_kv['EMAIL_LIDER']
            #mail.BCC = 'danieldelgadof@bcp.com.pe'
            
            fname = 'Formatos\AvanceCursos.html'
            html_file = open(fname, 'r',encoding='utf-8')
            html = html_file.read()
            mail = add_body(html, mail, input_tuNombre, dic_kv)
            print("Correo enviado 1")

        elif en == ETipoEnvio.InicioMedicion_0_porc_CC_lider: 
            
            porc = dic_kv['PORCENTAJE_AVANCE']
            if porc == 0:
                mail.CC=dic_kv['EMAIL_LIDER']
            fname = 'formatos\InicioMedicion.html'
            html_file = open(fname,'r', encoding='utf-8')
            html = html_file.read()
            mail = add_body(html, mail, input_tuNombre, dic_kv)
            print("Correo enviado 2")

        elif en == ETipoEnvio.InicioPDI:
            
            #mail.CC=dic_kv['EMAIL_LIDER']
            fname = 'Formatos\InicioPDI.html'
            html_file = open(fname,'r', encoding='utf-8')
            html = html_file.read()
            mail = add_body(html, mail, input_tuNombre, dic_kv)
            print("Correo enviado 3")
            
        elif en == ETipoEnvio.AvanceCurso_CC_lider:
            
            mail.CC=dic_kv['EMAIL_LIDER']
            fname = 'Formatos\AvanceCursos.html'
            html_file = open(fname,'r', encoding='utf-8')
            html = html_file.read()
            mail = add_body(html, mail, input_tuNombre, dic_kv)
            print("Correo enviado 4")
            
        elif en == ETipoEnvio.InicioMedicion_CC_lider:
            
            mail.CC=dic_kv['EMAIL_LIDER']
            fname = 'formatos\InicioMedicion.html'
            html_file = open(fname,'r', encoding='utf-8')
            html = html_file.read()
            mail = add_body(html, mail, input_tuNombre, dic_kv)
            print("Correo enviado 5")

        elif en == ETipoEnvio.Capability_building:
            
            #mail.CC=dic_kv['EMAIL_LIDER']
            fname = 'formatos\capability_building.html'
            html_file = open(fname,'r', encoding='utf-8')
            html = html_file.read()
            mail = add_body_with_image(html, mail, input_tuNombre, dic_kv)
            mail = add_files(mail,dic_kv)
            print("Correo enviado 6")

        elif en == ETipoEnvio.refuerzo_evaluacion:
            
            mail.CC=dic_kv['EMAIL_CAL'] #agregar mail de valeria, jenny y miluska bravo en el excel
            fname = 'formatos\Refuerzo_evaluacion.html'
            html_file = open(fname,'r', encoding='utf-8')
            html = html_file.read()
            mail = add_body_with_image(html, mail, input_tuNombre, dic_kv)
            print("Correo enviado 7")

        mail.Send()
        #mail.Save()

    
        #mail = add_files(mail, dic=dic)      
