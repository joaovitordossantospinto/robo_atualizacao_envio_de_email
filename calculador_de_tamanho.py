import os

def comparativo(anexo, kb_cabecalho):
    f_size = os.path.getsize(anexo)
    f_size_kb = f_size/1000
    if f_size_kb <= kb_cabecalho:
        flag_kb = 1
    else:
        flag_kb = 0
    return flag_kb    
