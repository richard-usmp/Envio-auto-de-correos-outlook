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
          Enviar Correo de:
          1. Avance de Curso con Lider en copia si el avance es 0%
          2. Inicio de Medición con Lider en copia si el avance es 0%
          3. Inicio de PDI
          4. Avance de Curso con Lider en copia
          5. Inicio de Medición con Lider en copia
          6. Capability building
          7. Refuerzo evaluación
          8. Finaliza tu autoevaluación con copia
          9. Autoevaluación finalizada
          10. Salir
          
          ELIJA UNA OPCIÓN: 
          ''')
    dic = {
        '1':ETipoEnvio.AvanceCurso_0_porc_CC_lider, 
        '2':ETipoEnvio.InicioMedicion_0_porc_CC_lider,
        '3':ETipoEnvio.InicioPDI,
        '4':ETipoEnvio.AvanceCurso_CC_lider,
        '5':ETipoEnvio.InicioMedicion_CC_lider,
        '6':ETipoEnvio.Capability_building,
        '7':ETipoEnvio.refuerzo_evaluacion,
        '8':ETipoEnvio.finaliza_autoevaluacion,
        '9':ETipoEnvio.autoevaluacion_finalizada
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
        emisor = input_emisor
        asunto = input_asunto 
        if en == ETipoEnvio.AvanceCurso_0_porc_CC_lider:  
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            send_emails(emisor, asunto, df, en, input_tuNombre)
            
        elif en == ETipoEnvio.InicioMedicion_0_porc_CC_lider:
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            #df['PRIMER NOMBRE'] = [input_tuPrimerNombre]
            send_emails(emisor, asunto, df, en, input_tuNombre)
            
        elif en == ETipoEnvio.InicioPDI:  
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            send_emails(emisor, asunto, df, en, input_tuNombre)
            
        elif en == ETipoEnvio.AvanceCurso_CC_lider:  
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            send_emails(emisor, asunto, df, en, input_tuNombre)
            
        elif en == ETipoEnvio.InicioMedicion_CC_lider: 
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            send_emails(emisor, asunto, df, en, input_tuNombre)

        elif en == ETipoEnvio.Capability_building: 
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            send_emails(emisor, asunto, df, en, input_tuNombre)

        elif en == ETipoEnvio.refuerzo_evaluacion: 
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            send_emails(emisor, asunto, df, en, input_tuNombre)

        elif en == ETipoEnvio.finaliza_autoevaluacion: 
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            #df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            send_emails(emisor, asunto, df, en, input_tuNombre)

        elif en == ETipoEnvio.autoevaluacion_finalizada: 
            df = pd.read_excel('excel_input_datos_formato\Formato envio correos.xlsx', sheet_name="BASE")
            df['S_PORCENTAJE_AVANCE'] = ['{:.0%}'.format(x) for x in df['PORCENTAJE_AVANCE'].values]
            send_emails(emisor, asunto, df, en, input_tuNombre)
            
    
if __name__ == '__main__':
    main()