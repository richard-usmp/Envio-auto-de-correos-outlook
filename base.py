import xlwings as xw
import pandas as pd
import numpy as np
from os import listdir, path

def send_email(email_emisor, email_receptor, asunto, dic=None):
    import win32com.client as win32
    ol=win32.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    mail=ol.CreateItem(olmailitem)
    mail.SentOnBehalfOfName = email_emisor
    mail.Subject= asunto
    mail.To=email_receptor
    return mail
    

def lastRow(ws, col=1):
    lwr_r_cell = ws.cells.last_cell
    lwr_row = lwr_r_cell.row
    lwr_cell = ws.range((lwr_row, col))

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('up')

    return lwr_cell.row

def lastColumn(ws, row=1):
    lwr_r_cell = ws.cells.last_cell
    lwr_col = lwr_r_cell.column
    lwr_cell = ws.range((row, lwr_col))

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('left')

    return lwr_cell.column

def leer_excel_simple(ruta,hoja=None,f_inicio=1, c_inicio=1,is_encuesta=False):
    header = 1

    app = xw.App(visible= False)
    app.display_alerts = False
    wb_api = app.books.api.Open(ruta, UpdateLinks=False, ReadOnly=True)
    wb = xw.Book(impl=xw._xlwindows.Book(xl=wb_api))
    
    ws = wb.sheets[0] if hoja is None else wb.sheets(hoja)
    # Obteneiendo rangos
    lr = lastRow(ws,c_inicio)
    lc = lastColumn(ws,f_inicio)

    # Caso encuesta
    if is_encuesta:
        header = 2 

    df = ws.range((f_inicio,c_inicio),(lr,lc)).options(pd.DataFrame, index=False,empty=np.nan, header=header).value

    wb.close()
    app.kill()

    return df

def add_body(html, mail, tu_nombre, dic=None):      
    if dic is not None:
        for k,v in dic.items():
            html = html.replace('$'+k,str(v))
            
    html = html.replace('$TU_NOMBRE', tu_nombre)
    mail.HTMLBody = html
    return mail

def add_body_with_image(html, mail, tu_nombre, dic=None):  
    lst_images = listdir('Imagenes')
    for i in lst_images:
       fn = str(i)
       r_path = path.join('Imagenes', fn)
       a_path = path.abspath(r_path)
       f1_at = mail.Attachments.Add(a_path)
       f1_at.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", fn)
       html = html.replace('Imagenes/'+i,'cid:'+fn)
        
    if dic is not None:
        for k,v in dic.items():
            html = html.replace('$'+k,str(v))

    html = html.replace('$TU_NOMBRE', tu_nombre)
    mail.HTMLBody = html
    return mail

def add_files(mail,dic=None,file_name=None,path_dir=None):
    p = f'Archivos' if path_dir is None else path_dir
    lst_files = listdir(p)

    for f in lst_files:
        fn = str(f)
        r_path = path.join('Archivos', fn)
        a_path = path.abspath(r_path)
        f1_at = mail.Attachments.Add(a_path)

    return mail