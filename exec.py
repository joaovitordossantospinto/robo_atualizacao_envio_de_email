import pyodbc
import pandas as pd
import os
import IPython
from acessos import dados_sql

import atualizador_de_arquivo
import calculador_de_tamanho
import disparador_de_email

server = dados_sql.get('server')
database = dados_sql.get('database')
username = dados_sql.get('username')
password = dados_sql.get('password')
pyodbc.pooling = False
cnn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnn.cursor()

sql_du = """SELECT DATA_BASE, DU
from DB_CALENDARIO_MIS
WHERE DATA_BASE = CAST(GETDATE() AS DATE)"""

df_du = pd.read_sql(sql_du,cnn)

sql = """       
DECLARE @HOJE DATE
        SET @HOJE = GETDATE()
SELECT DISTINCT
B.DATA_PROCESSAMENTO,A.ID_RELATORIO, A.NOME_RELATORIO, A.DIRETORIO_ARQUIVO, A.NOME_ARQUIVO
,A.EXTENSAO_ARQUIVO, A.ASSUNTO_EMAIL, A.DESTINATARIO_EMAIL, A.COPIA_EMAIL, A.NOME_DESTINATARIO_CORPO_EMAIL
,A.CORPO_EMAIL, A.FLAG_TEM_FECHAMENTO, A.KB_CABECALHO, C.STATUS, D.DU, A.FLAG_ANEXO, A.TIPO_ATUALIZACAO 
FROM RELATORIOS_AUTO A
INNER JOIN NOTIFICACAO_RELATORIOS_AUTO B ON B.ID_RELATORIO = A.ID_RELATORIO AND B.DATA_PROCESSAMENTO = @HOJE
LEFT JOIN STATUS_RELATORIOS_AUTO C ON C.ID_RELATORIO = A.ID_RELATORIO AND C.DATA = @HOJE
INNER JOIN DB_CALENDARIO_MIS D ON D.DATA_BASE = @HOJE
WHERE CAST(A.DIAS_SEMANA AS varchar) LIKE '%' + CAST(DATEPART(WEEKDAY, @HOJE) AS varchar) + '%'
AND A.FLAG_HABILITADO = 1
AND ISNULL(C.STATUS,'') NOT IN ('CONCLUIDO')
"""

df = pd.read_sql(sql,cnn)
    
if len(df_du) > 0:
    if df_du.at[0, 'DU'] == 1: 
        sql_ds = """
            DECLARE @HOJE DATE
                 SET @HOJE = GETDATE()
        SELECT DISTINCT
        B.DATA_PROCESSAMENTO,A.ID_RELATORIO, A.NOME_RELATORIO, A.DIRETORIO_ARQUIVO, A.NOME_ARQUIVO
        ,A.EXTENSAO_ARQUIVO, A.ASSUNTO_EMAIL, A.DESTINATARIO_EMAIL, A.COPIA_EMAIL, A.NOME_DESTINATARIO_CORPO_EMAIL
        ,A.CORPO_EMAIL, A.FLAG_TEM_FECHAMENTO, A.KB_CABECALHO, C.STATUS, D.DU, A.FLAG_ANEXO, A.TIPO_ATUALIZACAO 
        FROM RELATORIOS_AUTO A
        INNER JOIN NOTIFICACAO_RELATORIOS_AUTO B ON B.ID_RELATORIO = A.ID_RELATORIO AND B.DATA_PROCESSAMENTO = @HOJE
        LEFT JOIN STATUS_RELATORIOS_AUTO C ON C.ID_RELATORIO = A.ID_RELATORIO AND C.DATA = @HOJE
        INNER JOIN DB_CALENDARIO_MIS D ON D.DATA_BASE = @HOJE
        WHERE A.FLAG_HABILITADO = 1
        AND (FLAG_TEM_FECHAMENTO = 1 or CAST(A.DIAS_SEMANA AS varchar) LIKE '%' + CAST(DATEPART(WEEKDAY, @HOJE) AS varchar) + '%')
        AND ISNULL(C.STATUS,'') NOT IN ('CONCLUIDO')
        """
        df_ds = pd.read_sql(sql_ds,cnn)
        if len(df_ds) > 0:
            for index, row in df_ds.iterrows():
                id_relatorio = int(r'{}'.format(row.ID_RELATORIO))
                nome_relatorio = r'{}'.format(row.NOME_RELATORIO)
                diretorio_arquivo = r'{}'.format(row.DIRETORIO_ARQUIVO)
                nome_arquivo = r'{}'.format(row.NOME_ARQUIVO)
                extensao_arquivo = r'{}'.format(row.EXTENSAO_ARQUIVO)
                assunto_email = r'{}'.format(row.ASSUNTO_EMAIL)
                destinatario_email = r'{}'.format(row.DESTINATARIO_EMAIL)
                copia_email = r'{}'.format(row.COPIA_EMAIL)
                nome_destinatario_corpo_email = r'{}'.format(row.NOME_DESTINATARIO_CORPO_EMAIL)
                corpo_email = r'{}'.format(row.CORPO_EMAIL)
                flag_tem_fechamento = int(r'{}'.format(row.FLAG_TEM_FECHAMENTO))
                kb_cabecalho = float(r'{}'.format(row.KB_CABECALHO))
                flag_anexo = int(r'{}'.format(row.FLAG_ANEXO))
                tipo_atualizacao = int(r'{}'.format(row.TIPO_ATUALIZACAO))                
                if tipo_atualizacao == 1:
                    if flag_tem_fechamento == 1:
                        sql_nomenclatura_mes = """select case when month( EOMONTH(GETDATE(),-1)) = 1 then 'JANEIRO'
                        when month( EOMONTH(GETDATE(),-1)) = 2 then 'FEVEREIRO'
                        when month( EOMONTH(GETDATE(),-1)) = 3 then 'MARÇO'
                        when month( EOMONTH(GETDATE(),-1)) = 4 then 'ABRIL'
                        when month( EOMONTH(GETDATE(),-1)) = 5 then 'MAIO'
                        when month( EOMONTH(GETDATE(),-1)) = 6 then 'JUNHO'
                        when month( EOMONTH(GETDATE(),-1)) = 7 then 'JULHO'
                        when month( EOMONTH(GETDATE(),-1)) = 8 then 'AGOSTO'
                        when month( EOMONTH(GETDATE(),-1)) = 9 then 'SETEMBRO'
                        when month( EOMONTH(GETDATE(),-1)) = 10 then 'OUTRUBRO'
                        when month( EOMONTH(GETDATE(),-1)) = 11 then 'NOVEMBRO'
                        when month( EOMONTH(GETDATE(),-1)) = 12 then 'DEZEMBRO' ELSE '' END AS MES"""
                        df_nomenclatura_mes = pd.read_sql(sql_nomenclatura_mes,cnn)
                        diretorio_historico = atualizador_de_arquivo.atualizar_relatorio_fechamento(diretorio_arquivo, nome_arquivo, extensao_arquivo, df_nomenclatura_mes.at[0, 'MES'])
                        r_kb = calculador_de_tamanho.comparativo(diretorio_historico, kb_cabecalho)
                        if r_kb == 1:
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'RELATÓRIO POTENCIALEMNTE VAZIO')")               
                            cnn.commit()
                        else:
                            if flag_anexo == 1:
                                disparador_de_email.envia_email(nome_destinatario_corpo_email, diretorio_historico, destinatario_email, copia_email, '{} FECHAMENTO {}'.format(assunto_email, df_nomenclatura_mes.at[0, 'MES']), corpo_email )
                                cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                                cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                                cnn.commit()
                            else:
                                disparador_de_email.envia_email_sem_anexo(nome_destinatario_corpo_email, destinatario_email, copia_email, '{} FECHAMENTO {}'.format(assunto_email, df_nomenclatura_mes.at[0, 'MES']), corpo_email, diretorio_historico )
                                cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                                cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                                cnn.commit()
                    else:
                        diretorio_historico = atualizador_de_arquivo.atualizar_relatorio(r'{}\{}.{}'.format(diretorio_arquivo, nome_arquivo, extensao_arquivo))
                        r_kb = calculador_de_tamanho.comparativo(r'{}\{}.{}'.format(diretorio_arquivo, nome_arquivo, extensao_arquivo), kb_cabecalho)
                        if r_kb == 1:
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'RELATÓRIO POTENCIALEMNTE VAZIO')")               
                            cnn.commit()
                        else:
                            if flag_anexo == 1:
                                disparador_de_email.envia_email(nome_destinatario_corpo_email, diretorio_historico, destinatario_email, copia_email, assunto_email, corpo_email )
                                cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                                cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                                cnn.commit()
                            else:
                                disparador_de_email.envia_email_sem_anexo(nome_destinatario_corpo_email, destinatario_email, copia_email, assunto_email, corpo_email, diretorio_historico)
                                cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                                cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                                cnn.commit()
                elif tipo_atualizacao == 2:
                    diretorio_historico = atualizador_de_arquivo.atualizar_relatorio_data_diaria(diretorio_arquivo, nome_arquivo, extensao_arquivo)
                    r_kb = calculador_de_tamanho.comparativo(diretorio_historico, kb_cabecalho)
                    if r_kb == 1:
                        cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                        cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'RELATÓRIO POTENCIALEMNTE VAZIO')")               
                        cnn.commit()
                    else:
                        if flag_anexo == 1:
                            disparador_de_email.envia_email(nome_destinatario_corpo_email, diretorio_historico, destinatario_email, copia_email, assunto_email, corpo_email )
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                        else:
                            disparador_de_email.envia_email_sem_anexo(nome_destinatario_corpo_email, destinatario_email, copia_email, assunto_email, corpo_email, diretorio_historico)
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                elif tipo_atualizacao == 3:
                    diretorio_historico = atualizador_de_arquivo.atualizar_relatorio_data_diaria_sms_sicredi(diretorio_arquivo, nome_arquivo, extensao_arquivo)
                    r_kb = calculador_de_tamanho.comparativo(diretorio_historico, kb_cabecalho)
                    if r_kb == 1:
                        cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                        cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'RELATÓRIO POTENCIALEMNTE VAZIO')")               
                        cnn.commit()
                    else:
                        if flag_anexo == 1:
                            disparador_de_email.envia_email(nome_destinatario_corpo_email, diretorio_historico, destinatario_email, copia_email, assunto_email, corpo_email )
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                        else:
                            disparador_de_email.envia_email_sem_anexo(nome_destinatario_corpo_email, destinatario_email, copia_email, assunto_email, corpo_email, diretorio_historico)
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                elif tipo_atualizacao == 4:
                    diretorio_historico = atualizador_de_arquivo.atualizar_relatorio_data_diaria_sicredi_tempos(diretorio_arquivo, nome_arquivo, extensao_arquivo)
                    r_kb = calculador_de_tamanho.comparativo(diretorio_historico, kb_cabecalho)
                    if r_kb == 1:
                        cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                        cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'RELATÓRIO POTENCIALEMNTE VAZIO')")               
                        cnn.commit()
                    else:
                        if flag_anexo == 1:
                            disparador_de_email.envia_email(nome_destinatario_corpo_email, diretorio_historico, destinatario_email, copia_email, assunto_email, corpo_email )
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                        else:
                            disparador_de_email.envia_email_sem_anexo(nome_destinatario_corpo_email, destinatario_email, copia_email, assunto_email, corpo_email, diretorio_historico)
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()            
            sql_status_email = """DECLARE @HOJE DATE
                        SET @HOJE = GETDATE()
                        SELECT  D.DATA_BASE, A.ID_RELATORIO, A.NOME_RELATORIO, CASE WHEN B.STATUS IS NULL THEN 'AGUARDANDO CONCLUSÃO DE JOB' ELSE B.STATUS END AS STATUS
                        FROM RELATORIOS_AUTO A
                        LEFT JOIN STATUS_RELATORIOS_AUTO B ON B.ID_RELATORIO = A.ID_RELATORIO
                        AND B.DATA = CAST(GETDATE() AS DATE)
                        INNER JOIN DB_CALENDARIO_MIS D ON CAST(D.DATA_BASE AS DATE) = CAST(GETDATE() AS DATE)
                        WHERE A.FLAG_HABILITADO = 1
                        AND (FLAG_TEM_FECHAMENTO = 1 or CAST(A.DIAS_SEMANA AS varchar) LIKE '%' + CAST(DATEPART(WEEKDAY, @HOJE) AS varchar) + '%')
                        order by STATUS, ID_RELATORIO"""
            df_status_email = pd.read_sql(sql_status_email,cnn)
            df_semIndices = df_status_email.to_html(index=False)
            disparador_de_email.envia_email_status(df_semIndices)
        else:
            sql_status_email = """DECLARE @HOJE DATE
                            SET @HOJE = GETDATE()
                            SELECT  D.DATA_BASE, A.ID_RELATORIO, A.NOME_RELATORIO, CASE WHEN B.STATUS IS NULL THEN 'AGUARDANDO CONCLUSÃO DE JOB' ELSE B.STATUS END AS STATUS
                            FROM RELATORIOS_AUTO A
                            LEFT JOIN STATUS_RELATORIOS_AUTO B ON B.ID_RELATORIO = A.ID_RELATORIO
                            AND B.DATA = CAST(GETDATE() AS DATE)
                            INNER JOIN DB_CALENDARIO_MIS D ON CAST(D.DATA_BASE AS DATE) = CAST(GETDATE() AS DATE)
                            WHERE A.FLAG_HABILITADO = 1
                            AND (FLAG_TEM_FECHAMENTO = 1 or CAST(A.DIAS_SEMANA AS varchar) LIKE '%' + CAST(DATEPART(WEEKDAY, @HOJE) AS varchar) + '%')
                            order by STATUS, ID_RELATORIO"""
            df_status_email = pd.read_sql(sql_status_email,cnn)
            df_semIndices = df_status_email.to_html(index=False)
            disparador_de_email.envia_email_status(df_semIndices)
    else:
        if len(df) > 0:
            for index, row in df.iterrows():
                id_relatorio = int(r'{}'.format(row.ID_RELATORIO))
                nome_relatorio = r'{}'.format(row.NOME_RELATORIO)
                diretorio_arquivo = r'{}'.format(row.DIRETORIO_ARQUIVO)
                nome_arquivo = r'{}'.format(row.NOME_ARQUIVO)
                extensao_arquivo = r'{}'.format(row.EXTENSAO_ARQUIVO)
                assunto_email = r'{}'.format(row.ASSUNTO_EMAIL)
                destinatario_email = r'{}'.format(row.DESTINATARIO_EMAIL)
                copia_email = r'{}'.format(row.COPIA_EMAIL)
                nome_destinatario_corpo_email = r'{}'.format(row.NOME_DESTINATARIO_CORPO_EMAIL)
                corpo_email = r'{}'.format(row.CORPO_EMAIL)
                kb_cabecalho = float(r'{}'.format(row.KB_CABECALHO))
                flag_anexo = int(r'{}'.format(row.FLAG_ANEXO))
                tipo_atualizacao = int(r'{}'.format(row.TIPO_ATUALIZACAO))
                if tipo_atualizacao == 1:
                    atualizador_de_arquivo.atualizar_relatorio(r'{}\{}.{}'.format(diretorio_arquivo, nome_arquivo, extensao_arquivo))
                    r_kb = calculador_de_tamanho.comparativo(r'{}\{}.{}'.format(diretorio_arquivo, nome_arquivo, extensao_arquivo), kb_cabecalho)
                    if r_kb == 1:
                        cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                        cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'RELATÓRIO POTENCIALEMNTE VAZIO')")               
                        cnn.commit()
                    else:
                        if flag_anexo == 1:
                            disparador_de_email.envia_email(nome_destinatario_corpo_email, r'{}\{}.{}'.format(diretorio_arquivo, nome_arquivo, extensao_arquivo), destinatario_email, copia_email, assunto_email, corpo_email )
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                        else:
                            disparador_de_email.envia_email_sem_anexo(nome_destinatario_corpo_email, destinatario_email, copia_email, assunto_email, corpo_email, r'{}\{}.{}'.format(diretorio_arquivo, nome_arquivo, extensao_arquivo))
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                elif tipo_atualizacao == 2:
                    diretorio_historico = atualizador_de_arquivo.atualizar_relatorio_data_diaria(diretorio_arquivo, nome_arquivo, extensao_arquivo)
                    r_kb = calculador_de_tamanho.comparativo(diretorio_historico, kb_cabecalho)
                    if r_kb == 1:
                        cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                        cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'RELATÓRIO POTENCIALEMNTE VAZIO')")               
                        cnn.commit()
                    else:
                        if flag_anexo == 1:
                            disparador_de_email.envia_email(nome_destinatario_corpo_email, diretorio_historico, destinatario_email, copia_email, assunto_email, corpo_email )
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                        else:
                            disparador_de_email.envia_email_sem_anexo(nome_destinatario_corpo_email, destinatario_email, copia_email, assunto_email, corpo_email, diretorio_historico)
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                elif tipo_atualizacao == 3:
                    diretorio_historico = atualizador_de_arquivo.atualizar_relatorio_data_diaria_sms_sicredi(diretorio_arquivo, nome_arquivo, extensao_arquivo)
                    r_kb = calculador_de_tamanho.comparativo(diretorio_historico, kb_cabecalho)
                    if r_kb == 1:
                        cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                        cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'RELATÓRIO POTENCIALEMNTE VAZIO')")               
                        cnn.commit()
                    else:
                        if flag_anexo == 1:
                            disparador_de_email.envia_email(nome_destinatario_corpo_email, diretorio_historico, destinatario_email, copia_email, assunto_email, corpo_email )
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                        else:
                            disparador_de_email.envia_email_sem_anexo(nome_destinatario_corpo_email, destinatario_email, copia_email, assunto_email, corpo_email, diretorio_historico)
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                elif tipo_atualizacao == 4:
                    diretorio_historico = atualizador_de_arquivo.atualizar_relatorio_data_diaria_sicredi_tempos(diretorio_arquivo, nome_arquivo, extensao_arquivo)
                    r_kb = calculador_de_tamanho.comparativo(diretorio_historico, kb_cabecalho)
                    if r_kb == 1:
                        cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                        cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'RELATÓRIO POTENCIALEMNTE VAZIO')")               
                        cnn.commit()
                    else:
                        if flag_anexo == 1:
                            disparador_de_email.envia_email(nome_destinatario_corpo_email, diretorio_historico, destinatario_email, copia_email, assunto_email, corpo_email )
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()
                        else:
                            disparador_de_email.envia_email_sem_anexo(nome_destinatario_corpo_email, destinatario_email, copia_email, assunto_email, corpo_email, diretorio_historico)
                            cursor.execute(f"DELETE FROM MIS.DBO.STATUS_RELATORIOS_AUTO WHERE ID_RELATORIO = {id_relatorio} AND DATA = CAST(GETDATE() AS DATE)")
                            cursor.execute(f"INSERT INTO MIS.DBO.STATUS_RELATORIOS_AUTO (DATA, ID_RELATORIO, NOME_RELATORIO, STATUS) VALUES (GETDATE(), {id_relatorio}, '{nome_relatorio}', 'CONCLUIDO')")               
                            cnn.commit()            
            sql_status_email = """SELECT  D.DATA_BASE, A.ID_RELATORIO, A.NOME_RELATORIO, CASE WHEN B.STATUS IS NULL THEN 'AGUARDANDO CONCLUSÃO DE JOB' ELSE B.STATUS END AS STATUS
                        FROM RELATORIOS_AUTO A
                        LEFT JOIN STATUS_RELATORIOS_AUTO B ON B.ID_RELATORIO = A.ID_RELATORIO
                        AND B.DATA = CAST(GETDATE() AS DATE)
                        INNER JOIN DB_CALENDARIO_MIS D ON CAST(D.DATA_BASE AS DATE) = CAST(GETDATE() AS DATE)
                        WHERE CAST(A.DIAS_SEMANA AS varchar) LIKE '%' + CAST(DATEPART(WEEKDAY, GETDATE()) AS varchar) + '%'
                        AND A.FLAG_HABILITADO = 1
                        ORDER BY STATUS, ID_RELATORIO"""
            df_status_email = pd.read_sql(sql_status_email,cnn)
            df_semIndices = df_status_email.to_html(index=False)
            disparador_de_email.envia_email_status(df_semIndices)        
        else:
              sql_status_email = """SELECT  D.DATA_BASE, A.ID_RELATORIO, A.NOME_RELATORIO, CASE WHEN B.STATUS IS NULL THEN 'AGUARDANDO CONCLUSÃO DE JOB' ELSE B.STATUS END AS STATUS
                          FROM RELATORIOS_AUTO A
                          LEFT JOIN STATUS_RELATORIOS_AUTO B ON B.ID_RELATORIO = A.ID_RELATORIO
                          AND B.DATA = CAST(GETDATE() AS DATE)
                          INNER JOIN DB_CALENDARIO_MIS D ON CAST(D.DATA_BASE AS DATE) = CAST(GETDATE() AS DATE)
                          WHERE CAST(A.DIAS_SEMANA AS varchar) LIKE '%' + CAST(DATEPART(WEEKDAY, GETDATE()) AS varchar) + '%'
                          AND A.FLAG_HABILITADO = 1
                          ORDER BY STATUS, ID_RELATORIO"""
              df_status_email = pd.read_sql(sql_status_email,cnn)
              df_semIndices = df_status_email.to_html(index=False)
              disparador_de_email.envia_email_status(df_semIndices)
else:
    disparador_de_email.envia_email_calendario_vazio()
    
raise SystemExit
       