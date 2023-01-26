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
    x_menu = input('''
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
    en = dic.get(x_menu)
    input_emisor = input('''
          Inserte su correo BCP y mantenga su aplicación de Outlook abierta:
          ''')
    
    input_asunto = input('''
          Inserte Asunto:
          ''')
    input_tuNombre = input('''
          Escriba su nombre completo:
          ''')
    #input_tuPrimerNombre = input('''
    #      Escriba su primer nombre:
    #      ''')
    
    if en is not None:
        if en == ETipoEnvio.AvanceCurso:
            emisor = input_emisor
            asunto = input_asunto    
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            send_emails(emisor, asunto, df, en, input_tuNombre)
        elif en == ETipoEnvio.InicioMedicion:
            emisor = input_emisor
            asunto = input_asunto   
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            #df['PRIMER NOMBRE'] = [input_tuPrimerNombre]
            send_emails(emisor, asunto, df, en, input_tuNombre)
        elif en == ETipoEnvio.InicioPDI:
            emisor = input_emisor
            asunto = input_asunto   
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            send_emails(emisor, asunto, df, en, input_tuNombre)
    
if __name__ == '__main__':
    main()