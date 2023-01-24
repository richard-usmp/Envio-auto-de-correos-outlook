from base import send_email

def send_emails(emisor,asunto,df):
    for receptor in df['EMAIL'].values:
        df_e = df[df['EMAIL']==receptor].T
        df_e.reset_index(level=0, inplace=True)
        df_e.columns = ['K','V']

        dic_kv = {}
        for i,r in df_e.iterrows():
            dic_kv[r['K']] = r['V'] 
        #mail = 
        send_email(emisor, receptor, asunto,dic_kv)
        #mail.Save()