import email, smtplib, ssl
import sys
import pandas as pd
import numpy as np
import os
import time
import csv
import shutil
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from string import Template
import credentials
from credentials import EMAIL_USER, EMAIL_PASSWORD

## STEP 1.0 Setting directories, vars and files.

DEFAULT_DIR = '/home/cleto/Desktop/E_mail_Automation/Python_Scripts/' ### Note: Please, set the default directory where the file main.py is before run
BASE_FILE_DIR = '../input_files/pending_base_file/'
clavesIVAFrontera = ["TACCMF", "OAXMED", "TIJHVI", "ENSSRL","REYSAN", "NLTHMA"]
REPORT_FILE_DIR = '../output_files/Reports/'
PROVIDERS_FILE_DIR = '../input_files/providers/'
E1P2_INPUT_FILES_DIR = '../output_files/Submitted/E1P2_Submitted_files/'
TABLES_DIR = '../output_files/Tables/'
SUBMITTED_BASE_FILE = '../output_files/Submitted/E1P1_base_file/'

def Turning_into_HTML ( df ) :
    Set_claves = list(set(list(df['Clave'])))
    for clave in Set_claves:
        table_by_clave = df[df.Clave.eq(clave)]
        Set_etapas = list(set(list(table_by_clave['Etapa'])))
        for etapa in Set_etapas:
            table_by_clave_and_set = table_by_clave[table_by_clave.Etapa.eq(etapa)]
            table_by_clave_and_set.to_html(buf = f'table_{clave}_{etapa}.html' ,index = False)

def Submitted_files_CSV ( clave, df ) :

    os.chdir(DEFAULT_DIR)
    os.chdir(E1P2_INPUT_FILES_DIR)
    tstamp = pk_prefix
    table_by_clave = df[df.Clave.eq(clave)]
    #MV_first_output = MV_first_output[['Clave','Folio','Lesionado','Etapa', 'Entrega','Tipo', 'Lesion', 'Subtotal', 'IVA', 'Total']]
    
    table_by_clave = table_by_clave[['Clave','RAZON SOCIAL','Subtotal','IVA', 'Total']]
    RZN_SOC = list(set(list(table_by_clave['RAZON SOCIAL'])))
    SBT = table_by_clave['Subtotal'].sum() #a = table_by_clave_2nd_output['Subtotal'].sum()
    IVA_T = table_by_clave['IVA'].sum()
    Total_T = table_by_clave['Total'].sum()
    
    dict_before_df = {'ID': [1] ,'Clave' : [clave], 'Razon Social' : [RZN_SOC[0]] , 'Subtotal' : [SBT], 'IVA' : [IVA_T], 'Total' : [Total_T], 'Estatus' : ["ENVIADO"]}
    table_by_clave_after = pd.DataFrame( dict_before_df )
    table_by_clave_after.to_csv(encoding = 'utf-8',path_or_buf = f'{clave}{tstamp}.csv' ,index = False)

    os.chdir(DEFAULT_DIR)

def Turning_into_CSV ( df ) :
    Set_claves = list(set(list(df['Clave'])))
    for clave in Set_claves:
        table_by_clave = df[df.Clave.eq(clave)]
        Set_etapas = list(set(list(table_by_clave['Etapa'])))
        for etapa in Set_etapas:
            table_by_clave_and_set = table_by_clave[table_by_clave.Etapa.eq(etapa)]
            table_by_clave_and_set.to_csv(encoding = 'utf-8',path_or_buf = f'table_{clave}_{etapa}.csv' ,index = False) #Especifica la ruta

def Sending_E_mail ( df ) :
    Set_claves = list(set(list(df['Clave'])))
    for clave in Set_claves:
        df_by_clave = df[df.Clave.eq(clave)]
        Set_etapas = list(set(list(df_by_clave['Etapa'])))
        for etapa in Set_etapas:
            df_by_clave_and_set = df_by_clave[df_by_clave.Etapa.eq(etapa)]
            RFC = list(set(list(df_by_clave_and_set['RFC'])))[0]
            RazonSocial = list(set(list(df_by_clave_and_set['RAZON SOCIAL'])))[0]
            importeFactura = str(df_by_clave_and_set['Total'].sum())
            totalAtenciones = str(df_by_clave_and_set['Ones'].sum())
            #MV_first_output = MV_first_output[['Clave','Folio','Lesionado','Etapa', 'Entrega','Tipo', 'Lesion', 'Subtotal', 'IVA', 'Total']]
            jigsaw_falling_into_place = df_by_clave_and_set [['Clave','Folio','Lesionado','Etapa', 'Entrega','Tipo', 'Lesion', 'Subtotal', 'IVA', 'Total']]
            html_table = jigsaw_falling_into_place.to_html(index = False, justify = 'left')
            subject_date = pk_prefix
            subject = f'{clave}_{etapa}_{subject_date}'
            fechaProceso = datetime.today().strftime("%d/ %m/ %Y")
            etapa_str = str(etapa)

            body =  """
                <html>        
                    <head> 
                        <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>     	
                        <style> 		
                            table{ 		  
                                border:1; 		
                                }  		
                            table th { 		  
                                color: #fff; 		  
                                background-color: blue; 		
                                } 
                        </style>    
                    </head>        
                    <body>            
                        <h1><span style="color:#353693; font-size:16pt;">$RazonSocial</span> </h1>
                        <h2><span style="color:#353693; font-size:16pt;">$RFC</span></h2>
                        <h4>Estimado Proveedor, </h4>    
                        <h4>Por medio del presente solicito su apoyo para el env&#237;o de una factura global por un total de <span style="color:#353693; font-size:16pt;"> $ $importeFactura MXN </span>, as&#237; como sus respectivos archivos XML y PDF para cubrir el pago de <span style="color:#353693; font-size:16pt;"> $totalAtenciones </span> atenciones Etapas <span style="color:#353693; font-size:16pt;">[ $etapa_str ]</span>, adjunto relaci&#243;n actualizada al d&#237;a <span style="color:#353693; font-size:16pt;"> $fechaProceso </span> con la finalidad de programar su pago.
                        <h4><span style='text-transform: uppercase;'>NOTA: SI EL METODO DE PAGO DE SU FACTURA ES PUE, DEBERA LLEGAR A MAS TARDAR EL DIA 25 DEL MES EN CURSO,  DE NO SER ASI TENDRIA QUE REFACTURARSE EN PPD Y GENERAR EL COMPLEMENTO DE PAGO DE LA MISMA. </span>   </h4>
                        $html_table
                        <p>
                            <h4>Favor de no modificar la relaci&#243;n que se les env&#237;a, si tiene alguna duda estamos a sus &#243;rdenes.</h4>
                            <h4>Le solicitamos anexar su factura (XML y PDF) dando respuesta a este mismo correo con la finalidad de poder procesas sus pagos de manera &#225;gil.</h4> 	
                            <h4>Quedamos a la espera de su factura, agradeciendo su tiempo y atenci&#243;n a la misma.</h4>  	
                        </p>  
                    </body>    
                </html>
                """
            
            body_f = Template(body).safe_substitute(RFC = RFC, RazonSocial = RazonSocial, html_table = html_table,\
                        importeFactura = importeFactura, totalAtenciones = totalAtenciones,\
                        etapa_str = etapa_str, fechaProceso = fechaProceso)

            recipients_email_raw = list(set(list(df_by_clave_and_set['MAIL SOLICITUD DE FACTURAS'])))
            recipients_email_list = list(recipients_email_raw[0].split(", "))
            recipients_email = ", ".join(recipients_email_list)
            
            #print(recipients_email_list)
            #print(recipients_email)

            message = MIMEMultipart()
            message["From"] = EMAIL_USER
            message["To"] = recipients_email
            message["Subject"] = subject
            message.attach(MIMEText(body_f, "html"))

            filename = f'table_{clave}_{etapa}.csv'

            # Open the file of the table and code it.
            with open(filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

                encoders.encode_base64(part)
                # Add the name of attached file.
                    
                part.add_header("Content-Disposition",
                f"attachment; filename = {subject}.csv")

                # Attach the table
                message.attach(part)
                text_message = message.as_string()
    
                # Enter to the server
                smtp_ssl_host = 'mail.genericserver.com' #Your server here
                smtp_ssl_port = 465                      #Your port here
      
            try:
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL('mail.genericserver.com:465') #Your server and your port here
                server.login(EMAIL_USER, EMAIL_PASSWORD)
                print(f'\nMessage is prepared to send')
                server.sendmail(EMAIL_USER, recipients_email_list, text_message)
                server.quit()
                print(f'\nThe message has already sent to the following client: {clave}, stage: {etapa}, recipients: {recipients_email_list} ')
                time.sleep(15) #This pause is important to sent all the messages
            except Exception as Ex:
                print(f'ERROR: {Ex}') #Handle the exception.
        
        Submitted_files_CSV ( clave = clave , df = df )
        os.chdir(os.path.join(TABLES_DIR,pk_for_files))

def run():
    ## STEP 1.1 Setting base file.
    global pk_for_files
    global pk_prefix
    os.chdir(DEFAULT_DIR)
    pd.set_option('display.float_format', '{:,.2f}'.format)     #This line set the format of float numbers.
    pk_prefix = datetime.now().strftime("%d%m%Y")
    pk_for_files = pk_prefix + datetime.now().strftime("%H%M%S")     #This generate an unique primary key for files
    os.chdir(BASE_FILE_DIR)
    basefile_name = os.listdir() 
    df_basefile = pd.read_excel(basefile_name[0])
    df_basefile.rename( columns = {'Tabulador':'Subtotal'}, inplace = True)

    ##STEP 1.2 Joint both dataframes.

    df_basefile = df_basefile.set_index(['Clave','Folio'])
    #df_concfile = df_concfile.set_index(['Clave','Folio'])
    df_MV_first_output_raw = df_basefile #pd.concat([df_1, df_2], axis = 1, join = 'inner') 
    df_MV_first_output_raw = df_MV_first_output_raw.reset_index( drop = False)

    ##STEP 2.0 Create the first output file.

    df_MV_first_output = df_MV_first_output_raw[['Clave','Unidad','Folio','Lesionado','Etapa', 'Entrega','Tipo', 'Lesion', 'Subtotal', 'Consecutivo']]
    #df_MV_first_output[ 'Bool_Status' ] = df_MV_first_output[ 'Estatus' ].apply( lambda x: True if x == 'PENDIENTE' else False)
    
    ## STEP 2.1 Selecting just the pending services, defining IVA rates and operating over first dataframe.

    df_MV_first_output = df_MV_first_output.reset_index( drop = True)
    df_MV_first_output['IVA_rate'] = df_MV_first_output[ 'Clave' ].apply( lambda x: 0.08 if x in clavesIVAFrontera  else 0.16)
    df_MV_first_output['IVA'] = df_MV_first_output['Subtotal'] * df_MV_first_output['IVA_rate']
    df_MV_first_output['Total'] = df_MV_first_output['Subtotal'] + df_MV_first_output['IVA']
    df_MV_first_output['Ones'] = 1                         
    df_MV_first_output['Clave_Interna'] = df_MV_first_output.agg(lambda x: f"{x['Clave']}{pk_prefix}", axis=1)
    df_MV_first_output['Rel_Clave'] = df_MV_first_output.agg(lambda x: f"{x['Clave']}_{x['Consecutivo']}", axis=1)
    
    ## STEP 2.2 Formating the main dataframe. Create the report.

    df_MV_first_output = df_MV_first_output[['Clave','Unidad','Folio','Lesionado','Etapa', 'Entrega','Tipo', 'Lesion', 'Subtotal', 'IVA', 'Total', 'Consecutivo', 'Ones','Clave_Interna','Rel_Clave']]
    
    df_VM_first_output_for_report = df_MV_first_output[['Clave','Unidad','Folio','Lesionado','Etapa', 'Entrega','Tipo', 'Lesion', 'Subtotal', 'IVA', 'Total', 'Consecutivo','Clave_Interna','Rel_Clave']]
    os.chdir(DEFAULT_DIR)
    os.chdir(REPORT_FILE_DIR)
    df_VM_first_output_for_report.to_excel(f'Report_{pk_for_files}.xlsx', index = False)

    ## STEP 3.0 Create the second output file (with invoicing data and email addresses)
    os.chdir(DEFAULT_DIR)
    os.chdir(PROVIDERS_FILE_DIR)
    df_providers_dir = pd.read_excel( 'BASEPROVEEDORES_Test.xlsx') ### NOTE: Please, set this file name before testing eviroment.
    
    df_1 = df_MV_first_output.set_index(['Clave'])
    #print(df_1)
    df_1 = df_1.loc[~df_1.index.duplicated(keep='first')] # In order to prevent a shaping bug
    #print(df_1)
    df_2 = df_providers_dir.set_index(['Clave'])
    df_MV_second_output_raw = pd.concat([df_1, df_2], axis = 1, join = 'inner')
    df_MV_second_output_raw = df_MV_second_output_raw.reset_index( drop = False)
    df_MV_second_output = df_MV_second_output_raw[['Clave','Unidad','Folio','Lesionado','Etapa','Entrega','Tipo', 'Lesion', 'Subtotal', 'IVA', 'Total', 'Consecutivo','Ones','RAZON SOCIAL', 'RFC', 'MAIL SOLICITUD DE FACTURAS' ]]
    print(df_MV_second_output)

    ## STEP 4.0 Create tables and files before sending messages.
    os.chdir(DEFAULT_DIR)
    os.mkdir(os.path.join(TABLES_DIR,pk_for_files)) #https://www.geeksforgeeks.org/create-a-directory-in-python/
    os.chdir(os.path.join(TABLES_DIR,pk_for_files))
    Turning_into_HTML( df_MV_first_output )
    Turning_into_CSV ( df_MV_first_output )

    ## STEP 5.0 Sending e-mail.

    #print(df_MV_second_output)
    Sending_E_mail( df_MV_second_output )

    ## STEP 6.0 Move to submited and make the final outputs.

    os.chdir(DEFAULT_DIR)
    shutil.move(f"{BASE_FILE_DIR}{basefile_name[0]}",f"{SUBMITTED_BASE_FILE}{basefile_name[0]}")
    os.chdir(DEFAULT_DIR)

if __name__ == '__main__':
    run()