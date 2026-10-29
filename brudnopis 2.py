import pandas as pd

plik  = pd.read_csv(r'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\job_SO5IcVfVB7Q3wurCQzlMJQpgkzvB.csv')
df = pd.DataFrame(plik)
pd.set_option('display.max_columns', None)
pd.set_option('display.width',500)
df.index= range(1,len(df)+1)
df = df[['active_sold_to','active_bill_to','bill_to_mail','bill_to_mail_notes']]
df = df[df['bill_to_mail_notes'] == 'EDOC_DE']
print(df)