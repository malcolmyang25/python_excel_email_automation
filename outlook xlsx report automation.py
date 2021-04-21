import xlsxwriter
import pyodbc
import time
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

#initial environment
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=FFD-SQL01;DATABASE=XXX;UID=XXX;PWD=XXX')
nowday=time.strftime("%Y-%m-%d %H:%M",time.localtime(time.time()))
xls='XXX\\Desktop\\welldone report test\\Welocme1.xlsx

#load data from database
sql_query="SELECT * FROM vw_property_welldone"
df=pd.read_sql(sql_query,cnxn)

#cook excel sheet
workbook = xlsxwriter.Workbook(xls)
worksheet = workbook.add_worksheet()
cell_format = workbook.add_format()
cell_format.set_text_wrap()
cell_format.set_bold()
cell_list = ['A','B','C','D','E','F','G','H','I','J','K']
df_tittle = list(df.columns.values)
i=0
for i in range(len(df_tittle)):
    cell_loc=cell_list[i]+'1'
    cell_name=df_tittle[i]
    column_name=cell_list[i]
    column_wide = int(df[cell_name].str.encode(encoding='utf-8').str.len().max())
    worksheet.write(cell_loc, cell_name, cell_format)
    worksheet.set_column(column_name+':'+column_name,column_wide+5)
df_tittle = list(df.columns.values)
for i in range(len(df_tittle)):
    for n in range(len(df.index)):
        col_name=df_tittle[i]
        cell_value=df[col_name][n]
        cell_format = workbook.add_format()
        cell_color = df['Category'][n].strip().lower()
#         cell_format.set_pattern(1)
#         cell_format.set_fg_color(cell_color)
        cell_format.set_font_color(cell_color)
        worksheet.write(n+1, i, cell_value,cell_format)
workbook.close()

##attch outlook and send
file=xls
html = """<html>
  <head></head>
  <body>
    <p>Good morning All,<br><br>
       This is the lastest version of XXX report for XXX report attched sent from XXX team.<br><br>
       Please contact relevant XXX Officer to arrange a sign up.<br><br>
       Thanks and regards.
    </p>
  </body>
</html>
"""
part = MIMEText(html, 'html')
msg.attach(part)
attachment = MIMEApplication(open(file,'rb').read())
attachment.add_header('Content-Disposition', 'attachment', filename=fname)
msg.attach(attachment)
mailserver = smtplib.SMTP('smtp.office365.com',587)
mailserver.ehlo()
mailserver.starttls()
mailserver.login(FROM, PSW)
mailserver.sendmail(FROM, TO, msg.as_string())
mailserver.quit()