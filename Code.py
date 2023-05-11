import pandas as pd
import numpy as np
import pyodbc
import matplotlib.pyplot as plt
import warnings



import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders  
import datetime

warnings.filterwarnings('ignore')
cnxn_str = ("Driver={SQL Server Native Client 11.0};"
            "Server='servidor'"
            "Database='Base_de_dados';"
            "Trusted_Connection=yes;")
cnxn = pyodbc.connect(cnxn_str)
cursor = cnxn.cursor()

dados = pd.read_sql(
"""
SELECT 
MARA.ERSDA AS Data_Criação,
COUNT( MARA.ERSDA ) AS Quantidade
                                       
FROM MARA
WHERE MARA.ERSDA BETWEEN REPLACE(CONVERT(varchar,FORMAT(getdate(), 'yyyy-01-01')),'-','') AND REPLACE(CONVERT(varchar,FORMAT(getdate(), 'yyyy-12-31')),'-','') AND MARA.MTART = 'FERT'

GROUP BY MARA.ERSDA
ORDER BY Data_Criação        
""",cnxn )

dados["Data_Criação"] = pd.to_datetime(dados["Data_Criação"], dayfirst=True)
dados = dados.groupby(dados['Data_Criação'].dt.strftime('%B'), sort=False)['Quantidade'].sum().reset_index()
fig, ax = plt.subplots(figsize=(15, 7))
x = np.arange(len(dados.Data_Criação))
grafico_1_unidade = ax.bar(x = x, height=dados.Quantidade,  data=dados, color='Blue',edgecolor='black') 

data_day =  datetime.date.today()
day = data_day.day
month = data_day.month
year = data_day.year   

ax.set_title(f'Cadastro no ano {year} FERT', fontsize=12, pad=20, color='black')
#ax.set_xlabel('Data_Criação', fontsize=14, labelpad=10, color='black')
#ax.set_ylabel('Quantidade', fontsize=14, labelpad=10, color='black')
plt.tick_params(left = False, right = False,            bottom = False, labelleft = False, labelbottom = True )
ax.set_xticks(x)
ax.set_xticklabels(dados.Data_Criação, color='black', size=14)
# colocando o rótulo nas barras
grafico_1_unidade = ax.bar_label(grafico_1_unidade,  size=14, label_type="edge")
#fmt="%.01f"
#Legenda
#n = ax.legend(['Quantidade'], fontsize=15)
n1 = fig.savefig('Mêses_FERT_Cadastros.png', dpi=240, bbox_inches='tight')





import pandas as pd
import numpy as np
import pyodbc
import matplotlib.pyplot as plt
import warnings
import locale
try:

    locale.setlocale(locale.LC_ALL, 'pt_BR')
except:
    locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil')



warnings.filterwarnings('ignore')
cnxn_str = ("Driver={SQL Server Native Client 11.0};"
            "Server='servidor'"
            "Database='Base_de_dados';"
            "Trusted_Connection=yes;")
cnxn = pyodbc.connect(cnxn_str)
cursor = cnxn.cursor()

dados = pd.read_sql(



"""
SELECT 
MARA.ERSDA AS Data_Criação,
COUNT( MARA.ERSDA ) AS Quantidade
                                       
FROM MARA
WHERE MARA.ERSDA BETWEEN REPLACE(CONVERT(varchar,FORMAT(getdate(), 'yyyy-MM-01')),'-','') AND REPLACE(CONVERT(varchar,FORMAT(getdate(), 'yyyy-MM-31')),'-','') AND MARA.MTART = 'FERT'
GROUP BY MARA.ERSDA
ORDER BY Data_Criação        
""",cnxn )

dados["Data_Criação"] = pd.to_datetime(dados["Data_Criação"], dayfirst=True)
dados['Data_Criação'] = dados['Data_Criação'].dt.strftime('%d/%m')
fig, ax = plt.subplots(figsize=(15, 7))
x = np.arange(len(dados.Data_Criação))

grafico_1_unidade = ax.bar(x = x, height=dados.Quantidade,  data=dados, color='green',edgecolor='black') 

data_day =  datetime.date.today()
day = data_day.day
month = data_day.month
month_descr =  datetime.date.today().strftime('%B/%Y')

year = data_day.year   

ax.set_title(f' Quantidade Cadastro {month_descr} Cadastro', fontsize=14, pad=20, color='black')
#ax.set_xlabel('Data_Criação', fontsize=14, labelpad=10, color='black')
#ax.set_ylabel('Quantidade', fontsize=14, labelpad=10, color='black')
plt.tick_params(left = False, right = False,            bottom = False, labelleft = False, labelbottom = True )
ax.set_xticks(x)
ax.set_xticklabels(dados.Data_Criação, color='black')
# colocando o rótulo nas barras
grafico_1_unidade = ax.bar_label(grafico_1_unidade,  size=14, label_type="edge")
#fmt="%.01f"
#Legenda
#n = ax.legend(['Quantidade'], fontsize=15)

fig.savefig('Dia_2023_Cadastro.png', dpi=240, bbox_inches='tight')



import pandas as pd
import numpy as np
import pyodbc
import matplotlib.pyplot as plt
import warnings

import locale
try:

    locale.setlocale(locale.LC_ALL, 'pt_BR')
except:
    locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil')

warnings.filterwarnings('ignore')
cnxn_str = ("Driver={SQL Server Native Client 11.0};"
            "Server=cismssql03.ciser.com.br;"
            "Database=inteligcom;"
            "Trusted_Connection=yes;")
cnxn = pyodbc.connect(cnxn_str)
cursor = cnxn.cursor()

dados = pd.read_sql(
"""
SELECT 
MARA.MTART AS Tipo,
COUNT( MARA.MTART ) AS Quantidade
                                       
FROM MARA
WHERE MARA.ERSDA > CONVERT(varchar, DATEPART(YEAR,GETDATE())) AND MARA.MTART = 'FERT'
GROUP BY MARA.MTART
      
""",cnxn )

fig, ax = plt.subplots(figsize=(10, 5))
x = np.arange(len(dados.Tipo))
grafico_1_unidade = ax.bar(x = x, height=dados.Quantidade,  data=dados, color='brown',edgecolor='black') 
ax.set_title(f'Total Cadastro {year}', fontsize=12, pad=20, color='black')
#ax.set_xlabel('Data_Criação', fontsize=14, labelpad=10, color='black')
#ax.set_ylabel('Quantidade', fontsize=14, labelpad=10, color='black')
plt.tick_params(left = False, right = False,            bottom = False, labelleft = False, labelbottom = True )
ax.set_xticks(x)
ax.set_xticklabels(dados.Tipo, color='black')
# colocando o rótulo nas barras
grafico_1_unidade = ax.bar_label(grafico_1_unidade,  size=35, padding=-100, label_type="edge")
#fmt="%.01f"
#Legenda
#n = ax.legend(['Quantidade'], fontsize=15)

n3 = fig.savefig('Total_Fert.png', dpi=240 , bbox_inches='tight')



import pandas as pd
import numpy as np
import pyodbc
import matplotlib.pyplot as plt
import warnings



import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders  
import datetime

warnings.filterwarnings('ignore')
cnxn_str = ("Driver={SQL Server Native Client 11.0};"
            "Server='servidor'"
            "Database='Base_de_dados';"
            "Trusted_Connection=yes;")
cnxn = pyodbc.connect(cnxn_str)
cursor = cnxn.cursor()

dados = pd.read_sql(
"""
SELECT 
LEFT(MARA.ERSDA, 4) AS Data_Criação,
COUNT( MARA.ERSDA ) AS Quantidade
                                       
FROM MARA
WHERE MARA.MTART = 'FERT'
GROUP BY LEFT(MARA.ERSDA, 4)
ORDER BY Data_Criação      
""",cnxn )




fig, ax = plt.subplots(figsize=(18, 7))
ax.set_title('Todos anos FERT', fontsize=12, pad=20, color='black')
x = np.arange(len(dados.Data_Criação))


grafico_1_unidade = ax.bar(x = x, height=dados.Quantidade,  data=dados, color='gray',edgecolor='black') 

plt.tick_params(left=False, right=False, bottom=False, labelright=False  ,labelleft=False, labelbottom = True )
ax.set_xticks(x)


ax.set_xticklabels(dados.Data_Criação, color='black', rotation=70 ,size=12)


grafico_1_unidade = ax.bar_label(grafico_1_unidade, size=14, label_type="edge")


ax2 = ax.twinx()
ax2.plot(x, dados.Quantidade, color='red', linestyle='dashed'   )
ax2.tick_params(axis='y',left=False, right=False, bottom=False, labelright=False  ,labelleft=False, labelbottom = False )
ax2.set_xticks(x)



n4 = fig.savefig('Todos_anos_fert.png', dpi=240, bbox_inches='tight')



# library
from PIL import Image
import matplotlib.pyplot as plt
  
# opening up of images
Dia_2023_Cadastro = Image.open("Dia_2023_Cadastro.png")
Dia_2023_Cadastro.size
Dia_2023_Cadastro = Dia_2023_Cadastro.resize((900,540))


Mêses_FERT_Cadastros = Image.open("Mêses_FERT_Cadastros.png")
Mêses_FERT_Cadastros.size
Mêses_FERT_Cadastros = Mêses_FERT_Cadastros.resize((900,540))



Total_Fert = Image.open("Total_Fert.png")
Total_Fert.size
Total_Fert = Total_Fert.resize((645,340))



Todos_anos_fert = Image.open("Todos_anos_fert.png")
Todos_anos_fert.size
Todos_anos_fert = Todos_anos_fert.resize((1100,540))



  

fundo = Image.new("RGB", (1940,1080), "white" )


fundo.paste(Dia_2023_Cadastro, (0, 0))
fundo.paste(Mêses_FERT_Cadastros, (900, 0))
fundo.paste(Total_Fert, (1150, 640))


fundo.paste(Todos_anos_fert, (0, 540))

fundo.save('fundo.png', bbox_inches='tight')  


import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders  
import datetime
data_day =  datetime.date.today()
day = data_day.day
month = data_day.month
year = data_day.year     
remetente = 'email'
senha_rede = 'senha**'
destinatario = 'email'
destinatario2 = 'email'
assunto = 'Relatório Quantidade de Cadastros'
# Preenche abaixo o corpo da mensagem.
texto = f"""
    Relatório Quantidade de Cadastros



    OBS: MENSAGEM AUTOMÁTICA.
"""
email_sender = remetente
msg = MIMEMultipart()
msg['From'] = remetente
msg['To'] = destinatario
#msg['To'] = destinatario2

msg['Subject'] = assunto

part = MIMEBase('application', "octet-stream")
imagem1 = 'fundo.png'
imagem2 = 'Mêses_2023_FERT_Cadastros.png'
part.set_payload(open("fundo.png", "rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="fundo.png"')
msg.attach(part)

msg.attach(MIMEText(_text=texto.encode('utf-8'), _charset='utf-8'))
port = port if '@empresa' in destinatario else 25
server = smtplib.SMTP(host='smtp.office365.com', port=port)
server.ehlo()
server.starttls()
server.login(remetente, senha_rede)
text = msg.as_string()
server.sendmail(email_sender, destinatario, text)
server.sendmail(email_sender, destinatario2, text)
print('Email enviado')
server.quit()
