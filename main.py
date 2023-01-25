import pandas as pd
from pcorreo import send_emails
from base import leer_excel_simple
from enumeraciones import ETipoEnvio

def importCSV():
    data = pd.read_csv("formato envio correos.csv", sep=";")   
    #mydata0 = pd.read_csv("workingfile.csv", skiprows=1, names=['CustID', 'Name', 'Companies', 'Income'])
    df = pd.DataFrame(data)
    return df.itertuples()

def main():
    x = input('''
          MENÚ
          ----------------------
          1. Avance de Curso
          2. Inicio de Medición
          3. Inicio de PDI
          4. Salir
          
          ELIJA UNA OPCIÓN: 
          ''')
    dic = {
        '1':ETipoEnvio.AvanceCurso, 
        '2':ETipoEnvio.InicioMedicion,
        '3':ETipoEnvio.InicioPDI
    }
    en = dic.get(x)
    
    if en is not None:
        if en == ETipoEnvio.AvanceCurso:
            emisor = 'andrepisco@bcp.com.pe'
            asunto = 'Correo de prueba - Seguimiento'    
            df = pd.read_excel('formato envio correos.xlsx', sheet_name="BASE")
            df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            send_emails(emisor, asunto,df)
        elif en == ETipoEnvio.InicioMedicion:
            emisor = 'andrepisco@bcp.com.pe'
            asunto = 'Correo de prueba - Seguimiento'    
            df = pd.read_excel('formato envio correos.xlsx', sheet_name="BASE")
            send_emails(emisor, asunto,df)
    
if __name__ == '__main__':
    main()
    # emisor = 'ricardoleon@bcp.com.pe'
    # asunto = 'Correo de prueba - Seguimiento'    
    # df = pd.read_excel('formato envio correos.xlsx', sheet_name="BASE")
    # df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
    # send_emails(emisor, asunto,df)