import win32com.client as win32
import datetime
from pathlib import Path
import pyodbc
from acessos import dados_sql
import pandas as pd
import disparador_de_email

server = dados_sql.get('server')
database = dados_sql.get('database')
username = dados_sql.get('username')
password = dados_sql.get('password')
pyodbc.pooling = False
cnn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnn.cursor()


data_atual = datetime.datetime.today().strftime('%Y%m%d')


excel = win32.Dispatch('Excel.Application')
excel.Visible = True
tipo = 'atualizacao'

def atualizar_relatorio(anexo, tentativas=3): #quantidade de tentativas adcionada
    if tentativas == 0:
        print("Limite de tentativas alcançado. Não foi possível atualizar o relatório.")
        return    
    try:
        wb = excel.Workbooks.Open(anexo)
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close(SaveChanges=False)
        excel.Quit()
        return anexo
    except:
                
        import wmi 
        contador = 0      
        processo = 'EXCEL.EXE'
        f = wmi.WMI()   
        for process in f.Win32_Process():       
            if process.name == processo: 
              process.Terminate() 
              contador += 1
        if contador == 0: 
            print(f"Nenhum processo {processo} encerrado")
        else:
            print(f"{contador} processos {processo} encerrados")
          
        atualizar_relatorio(anexo, tentativas=tentativas-1)
    
def atualizar_relatorio_fechamento(diretorio_arquivo, nome_arquivo, extensao_arquivo, df):
    try:
        wb = excel.Workbooks.Open(r'{}\{}_FECHAMENTO.{}'.format(diretorio_arquivo, nome_arquivo, extensao_arquivo))
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        diretorio_historico = r'{}\OLD\{}_FECHAMENTO_{}.{}'.format(diretorio_arquivo, nome_arquivo, df, extensao_arquivo)
        wb.SaveAs (diretorio_historico)
        wb.Close(SaveChanges=False)
        excel.Quit()    
        return diretorio_historico
    except:
                
        import wmi 
        contador = 0      
        processo = 'EXCEL.EXE'
        f = wmi.WMI()   
        for process in f.Win32_Process():       
            if process.name == processo: 
              process.Terminate() 
              contador += 1
        if contador == 0: 
            print(f"Nenhum processo {processo} encerrado")
            atualizar_relatorio_fechamento(diretorio_arquivo, nome_arquivo, extensao_arquivo, df)
        else:
            print(f"{contador} processos {processo} encerrados")
            atualizar_relatorio_fechamento(diretorio_arquivo, nome_arquivo, extensao_arquivo, df)
          
def atualizar_relatorio_data_diaria(diretorio, arquivo, extensao):   
    try:
        data_modificacao = lambda f: f.stat().st_mtime
        directory = Path(diretorio)
        files = directory.glob(f'{arquivo}*.{extensao}')
        sorted_files = sorted(files, key=data_modificacao, reverse=True)
        xlapp = win32.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open("{}".format(sorted_files[0]))
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        diretorio_historico = "{}\{}_{}.{}".format(diretorio, arquivo ,data_atual, extensao)  
        wb.SaveAs (diretorio_historico)
        wb.Close(SaveChanges=False)
        xlapp.Quit()
        return diretorio_historico
    except:
                
        import wmi 
        contador = 0      
        processo = 'EXCEL.EXE'
        f = wmi.WMI()   
        for process in f.Win32_Process():       
            if process.name == processo: 
              process.Terminate() 
              contador += 1
        if contador == 0: 
            print(f"Nenhum processo {processo} encerrado")
            atualizar_relatorio_data_diaria_sms_sicredi(diretorio, arquivo, extensao)
        else:
            print(f"{contador} processos {processo} encerrados")
            atualizar_relatorio_data_diaria_sms_sicredi(diretorio, arquivo, extensao)
                

def atualizar_relatorio_data_diaria_sms_sicredi(diretorio, arquivo, extensao):
    try:
        numero_de_registros_tabela = "SELECT count(*)+1 as count FROM MIS_SICREDI_GERAL_ENVIO_SMS_DIARIO"
        df = pd.read_sql(numero_de_registros_tabela,cnn)
        df2 = df.at[0, 'count'].astype(int)
        print(df2)
        data_modificacao = lambda f: f.stat().st_mtime
        directory = Path(diretorio)
        files = directory.glob(f'{arquivo}*.{extensao}')
        sorted_files = sorted(files, key=data_modificacao, reverse=True)
        xlapp = win32.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open("{}".format(sorted_files[0]))
        planilha = wb.Worksheets['ARQUIVO_SMS']
        numero_de_registros_ac = planilha.UsedRange.Rows.Count
        print(numero_de_registros_ac)
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        numero_de_registros_dc = planilha.UsedRange.Rows.Count
        print(numero_de_registros_dc)
        if numero_de_registros_dc == df2 and numero_de_registros_dc != numero_de_registros_ac:
            disparador_de_email.envia_email_validacao_sms_sicredi(df2, numero_de_registros_ac, numero_de_registros_dc, 'green', 'CORRETO!')
        else:
            disparador_de_email.envia_email_validacao_sms_sicredi(df2, numero_de_registros_ac, numero_de_registros_dc, 'red', 'OPOSTO DE CORRETO!')
        diretorio_historico = "{}\{}_{}.{}".format(diretorio, arquivo ,data_atual, extensao)  
        wb.SaveAs (diretorio_historico)
        wb.Close(SaveChanges=False)
        xlapp.Quit()
        return diretorio_historico 
    except:
                
        import wmi 
        contador = 0      
        processo = 'EXCEL.EXE'
        f = wmi.WMI()   
        for process in f.Win32_Process():       
            if process.name == processo: 
              process.Terminate() 
              contador += 1
        if contador == 0: 
            print(f"Nenhum processo {processo} encerrado")
            atualizar_relatorio_data_diaria_sms_sicredi(diretorio, arquivo, extensao)
        else:
            print(f"{contador} processos {processo} encerrados")
            atualizar_relatorio_data_diaria_sms_sicredi(diretorio, arquivo, extensao)
            
def atualizar_relatorio_data_diaria_sicredi_tempos(diretorio, arquivo, extensao):   
    try:
        wb = excel.Workbooks.Open('{}\{}.{}'.format(diretorio, arquivo, extensao))
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close(SaveChanges=False)
        excel.Quit()
        df2 = pd.read_excel('{}\{}.{}'.format(diretorio, arquivo, extensao))
        diretorio_historico = r"{}\{}_{}.{}".format(diretorio, data_atual ,'Arquivo_Tempos_Pausas_JARezende' , 'csv')  
        df2.to_csv(diretorio_historico, index=False)
        return diretorio_historico
    except:
                
        import wmi 
        contador = 0      
        processo = 'EXCEL.EXE'
        f = wmi.WMI()   
        for process in f.Win32_Process():       
            if process.name == processo: 
              process.Terminate() 
              contador += 1
        if contador == 0: 
            print(f"Nenhum processo {processo} encerrado")
            atualizar_relatorio_data_diaria_sicredi_tempos(diretorio, arquivo, extensao)
        else:
            print(f"{contador} processos {processo} encerrados") 
            atualizar_relatorio_data_diaria_sicredi_tempos(diretorio, arquivo, extensao)
            