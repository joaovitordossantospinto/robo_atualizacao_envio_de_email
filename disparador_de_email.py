import win32com.client as win32
import datetime

outlook = win32.Dispatch('Outlook.application')
hora = datetime.datetime.today().strftime('%H:%M')
data = datetime.datetime.today().strftime('%d/%m/%Y - %H:%M:%S')
saudacao = ''
tipo = 'email'

if hora < '12:00':
    saudacao = 'Bom dia'
elif hora >= '12:00' and hora < '18:00' :
    saudacao = 'Boa tarde'      
else:
    saudacao = 'Boa noite'        

def envia_email(nome, anexo, destinatario, copia, assunto, corpo):
    email = outlook.CreateItem(0)
    email.To = destinatario
    email.CC = copia
    email.Subject = assunto
    a1 = '{behavior:url(#default#VML);}'
    a2 = '{font-family:"Cambria Math";ose-1:2 4 5 3 5 4 6 3 2 4;}'
    a3 = '{font-family:Lato;}'
    a4 = '{margin:0cm;t-size:11.0pt;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a5 = '{mso-style-type:personal-compose;t-family:"Calibri",sans-serif;or:windowtext;}'
    a6 = '{mso-style-type:export-only;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a7 = '{size:612.0pt 792.0pt;gin:70.85pt 3.0cm 70.85pt 3.0cm;}'
    a8 = '{page:WordSection1;}'
    img = r"\\192.168.200.14\JARezende\Administrativo\MIS\Planejamento_MIS\MIS\00-PESSOAL\JOAO\O_MAIOR_PROJETO_DE_TODOS_OS_TEMPOS\RELATORIOS_AUTO\img-assinatura.png"

    email.HTMLBody = f'''
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=iso-8859-1"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {a1}
    o\:* {a1}
    w\:* {a1}
    .shape {a1}
    </style><![endif]--><style><!--
    /* Font Definitions */
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a3}
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
    	{a4}
    span.EstiloDeEmail17
    	{a5}
    .MsoChpDefault
    	{a6}
    @page WordSection1
    	{a7}
    div.WordSection1
    	{a8}
    --></style><!--[if gte mso 9]><xml>
    <o:shapedefaults v:ext="edit" spidmax="1026" />
    </xml><![endif]--><!--[if gte mso 9]><xml>
    <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1" />
    </o:shapelayout></xml><![endif]-->
    </head>
    <body lang=PT-BR link="#0563C1" vlink="#954F72" style='word-wrap:break-word'>
    <div class=WordSection1>
    <p class=MsoNormal>
    <span style='font-size:12.0pt'>
    {nome}, {saudacao}!
    <br>
    <br>Segue anexo, {corpo}.
    <br>
    <br>Att.</span>
    <o:p>
    </o:p>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <span style='mso-fareast-language:PT-BR'>
    <o:p>&nbsp;</o:p>
    </span>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <b>
    <span style='font-size:12.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Relatórios Mis
    <br>
    </span>
    </b>
    <span style='font-size:10.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Envio automático 
    <br>
    <br>Bruno Abad Portugal Velasco -  Ramal 5288 / E-mail bruno.velasco@jarezende.com.br
    <br>Bruno de Almeida Frias Andriolli -  Ramal 5287 / E-mail bruno.andriolli@jarezende.com.br
    <br>Bruno Ezequiel Celidonio - Ramal 2734 / E-mail bruno.sousa@jarezende.com.br
    <br>João Vitor dos Santos Pinto - Ramal 2156 / E-mail joao.santos@jarezende.com.br
    <br>
    <br>
    <img src="{img}" width="517" height="117">
    <br>
    <br>
    </span><i>
    <span style='font-size:8.0pt;line-height:106%;font-family:"Calibri Light",sans-serif;color:gray;mso-fareast-language:PT-BR'>
    Esta mensagem pode conter informação confidencial e/ou privilegiada. 
    Se você não for o destinatário ou a pessoa autorizada a receber esta mensagem, não pode usar, 
    copiar ou divulgar as informações nela contidas ou tomar qualquer ação baseada nessas informações. 
    Se você recebeu esta mensagem por engano, por favor, avise imediatamente o remetente, respondendo o e-mail, 
    e em seguida apague-o. 
    Agradecemos sua cooperação.
    <br>
    <br>
    </span>
    </i>
    <i>
    <span style='font-size:8.0pt;line-height:106%;color:gray;mso-fareast-language:PT-BR'>
    This message may contain confidential and/or privileged information. 
    If you are not the address or authorized to receive this for the address, you must not use, copy, 
    disclose or take any action base on this message or any information herein. 
    If you have received this message in error, please advise the sender immediately by reply e-mail and delete this message. 
    Thank you for your cooperation<o:p>
    </o:p></span></i></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>
    '''

    email.Attachments.Add(anexo)
    email.Send()
    
def envia_email_status(df):
    email = outlook.CreateItem(0)
    email.To = 'emailficticiodaminhaequipe@gmail.com'
    # email.CC = ''
    email.Subject = 'STATUS RELATÓRIOS DIÁRIO {}'.format(data)
    a1 = '{behavior:url(#default#VML);}'
    a2 = '{font-family:"Cambria Math";ose-1:2 4 5 3 5 4 6 3 2 4;}'
    a3 = '{font-family:Lato;}'
    a4 = '{margin:0cm;t-size:11.0pt;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a5 = '{mso-style-type:personal-compose;t-family:"Calibri",sans-serif;or:windowtext;}'
    a6 = '{mso-style-type:export-only;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a7 = '{size:612.0pt 792.0pt;gin:70.85pt 3.0cm 70.85pt 3.0cm;}'
    a8 = '{page:WordSection1;}'
    img = r"\\192.168.200.14\JARezende\Administrativo\MIS\Planejamento_MIS\MIS\00-PESSOAL\JOAO\O_MAIOR_PROJETO_DE_TODOS_OS_TEMPOS\RELATORIOS_AUTO\img-assinatura.png"

    email.HTMLBody = f'''
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=iso-8859-1"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {a1}
    o\:* {a1}
    w\:* {a1}
    .shape {a1}
    </style><![endif]--><style><!--
    /* Font Definitions */
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a3}
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
    	{a4}
    span.EstiloDeEmail17
    	{a5}
    .MsoChpDefault
    	{a6}
    @page WordSection1
    	{a7}
    div.WordSection1
    	{a8}
    --></style><!--[if gte mso 9]><xml>
    <o:shapedefaults v:ext="edit" spidmax="1026" />
    </xml><![endif]--><!--[if gte mso 9]><xml>
    <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1" />
    </o:shapelayout></xml><![endif]-->
    </head>
    <body lang=PT-BR link="#0563C1" vlink="#954F72" style='word-wrap:break-word'>
    <div class=WordSection1>
    <p class=MsoNormal>
    <span style='font-size:12.0pt'>
    Prezados, {saudacao}!
    <br>
    <br>Segue status atual dos relatórios diários.
    <br>
    <br>{df}
    <br>Att.</span>
    <o:p>
    </o:p>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <span style='mso-fareast-language:PT-BR'>
    <o:p>&nbsp;</o:p>
    </span>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <b>
    <span style='font-size:12.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Relatórios Mis
    <br>
    </span>
    </b>
    <span style='font-size:10.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Envio automático 
    <br>
    <br>Bruno Abad Portugal Velasco -  Ramal 5288 / E-mail bruno.velasco@jarezende.com.br
    <br>Bruno de Almeida Frias Andriolli -  Ramal 5287 / E-mail bruno.andriolli@jarezende.com.br
    <br>Bruno Ezequiel Celidonio - Ramal 2734 / E-mail bruno.sousa@jarezende.com.br
    <br>João Vitor dos Santos Pinto - Ramal 2156 / E-mail joao.santos@jarezende.com.br
    <br>
    <br>
    <img src="{img}" width="517" height="117">
    <br>
    <br>
    </span><i>
    <span style='font-size:8.0pt;line-height:106%;font-family:"Calibri Light",sans-serif;color:gray;mso-fareast-language:PT-BR'>
    Esta mensagem pode conter informação confidencial e/ou privilegiada. 
    Se você não for o destinatário ou a pessoa autorizada a receber esta mensagem, não pode usar, 
    copiar ou divulgar as informações nela contidas ou tomar qualquer ação baseada nessas informações. 
    Se você recebeu esta mensagem por engano, por favor, avise imediatamente o remetente, respondendo o e-mail, 
    e em seguida apague-o. 
    Agradecemos sua cooperação.
    <br>
    <br>
    </span>
    </i>
    <i>
    <span style='font-size:8.0pt;line-height:106%;color:gray;mso-fareast-language:PT-BR'>
    This message may contain confidential and/or privileged information. 
    If you are not the address or authorized to receive this for the address, you must not use, copy, 
    disclose or take any action base on this message or any information herein. 
    If you have received this message in error, please advise the sender immediately by reply e-mail and delete this message. 
    Thank you for your cooperation<o:p>
    </o:p></span></i></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>
    '''

    email.Send()    
    
def envia_email_calendario_vazio():
    email = outlook.CreateItem(0)
    email.To = 'emailficticiodaminhaequipe@gmail.com'
    email.Subject = 'STATUS RELATÓRIOS DIÁRIO {}'.format(data)
    a1 = '{behavior:url(#default#VML);}'
    a2 = '{font-family:"Cambria Math";ose-1:2 4 5 3 5 4 6 3 2 4;}'
    a3 = '{font-family:Lato;}'
    a4 = '{margin:0cm;t-size:11.0pt;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a5 = '{mso-style-type:personal-compose;t-family:"Calibri",sans-serif;or:windowtext;}'
    a6 = '{mso-style-type:export-only;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a7 = '{size:612.0pt 792.0pt;gin:70.85pt 3.0cm 70.85pt 3.0cm;}'
    a8 = '{page:WordSection1;}'
    img = r"\\192.168.200.14\JARezende\Administrativo\MIS\Planejamento_MIS\MIS\00-PESSOAL\JOAO\O_MAIOR_PROJETO_DE_TODOS_OS_TEMPOS\RELATORIOS_AUTO\img-assinatura.png"

    email.HTMLBody = f'''
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=iso-8859-1"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {a1}
    o\:* {a1}
    w\:* {a1}
    .shape {a1}
    </style><![endif]--><style><!--
    /* Font Definitions */
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a3}
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
    	{a4}
    span.EstiloDeEmail17
    	{a5}
    .MsoChpDefault
    	{a6}
    @page WordSection1
    	{a7}
    div.WordSection1
    	{a8}
    --></style><!--[if gte mso 9]><xml>
    <o:shapedefaults v:ext="edit" spidmax="1026" />
    </xml><![endif]--><!--[if gte mso 9]><xml>
    <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1" />
    </o:shapelayout></xml><![endif]-->
    </head>
    <body lang=PT-BR link="#0563C1" vlink="#954F72" style='word-wrap:break-word'>
    <div class=WordSection1>
    <p class=MsoNormal>
    <span style='font-size:12.0pt'>
    Prezados, {saudacao}!
    <br>
    <br>TABELA DB_CALENDARIO_MIS POSSIVELMENTE VÁZIA.
    <br>
    <br>Att.</span>
    <o:p>
    </o:p>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <span style='mso-fareast-language:PT-BR'>
    <o:p>&nbsp;</o:p>
    </span>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <b>
    <span style='font-size:12.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Relatórios Mis
    <br>
    </span>
    </b>
    <span style='font-size:10.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Envio automático 
    <br>
    <br>Bruno Abad Portugal Velasco -  Ramal 5288 / E-mail bruno.velasco@jarezende.com.br
    <br>Bruno de Almeida Frias Andriolli -  Ramal 5287 / E-mail bruno.andriolli@jarezende.com.br
    <br>Bruno Ezequiel Celidonio - Ramal 2734 / E-mail bruno.sousa@jarezende.com.br
    <br>João Vitor dos Santos Pinto - Ramal 2156 / E-mail joao.santos@jarezende.com.br
    <br>
    <br>
    <img src="{img}" width="517" height="117">
    <br>
    <br>
    </span><i>
    <span style='font-size:8.0pt;line-height:106%;font-family:"Calibri Light",sans-serif;color:gray;mso-fareast-language:PT-BR'>
    Esta mensagem pode conter informação confidencial e/ou privilegiada. 
    Se você não for o destinatário ou a pessoa autorizada a receber esta mensagem, não pode usar, 
    copiar ou divulgar as informações nela contidas ou tomar qualquer ação baseada nessas informações. 
    Se você recebeu esta mensagem por engano, por favor, avise imediatamente o remetente, respondendo o e-mail, 
    e em seguida apague-o. 
    Agradecemos sua cooperação.
    <br>
    <br>
    </span>
    </i>
    <i>
    <span style='font-size:8.0pt;line-height:106%;color:gray;mso-fareast-language:PT-BR'>
    This message may contain confidential and/or privileged information. 
    If you are not the address or authorized to receive this for the address, you must not use, copy, 
    disclose or take any action base on this message or any information herein. 
    If you have received this message in error, please advise the sender immediately by reply e-mail and delete this message. 
    Thank you for your cooperation<o:p>
    </o:p></span></i></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>
    '''

    email.Send()     
    
def envia_email_sem_anexo(nome, destinatario, copia, assunto, corpo, complemento_corpo_email):
    email = outlook.CreateItem(0)
    email.To = destinatario
    email.CC = copia
    email.Subject = assunto
    a1 = '{behavior:url(#default#VML);}'
    a2 = '{font-family:"Cambria Math";ose-1:2 4 5 3 5 4 6 3 2 4;}'
    a3 = '{font-family:Lato;}'
    a4 = '{margin:0cm;t-size:11.0pt;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a5 = '{mso-style-type:personal-compose;t-family:"Calibri",sans-serif;or:windowtext;}'
    a6 = '{mso-style-type:export-only;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a7 = '{size:612.0pt 792.0pt;gin:70.85pt 3.0cm 70.85pt 3.0cm;}'
    a8 = '{page:WordSection1;}'
    img = r"\\192.168.200.14\JARezende\Administrativo\MIS\Planejamento_MIS\MIS\00-PESSOAL\JOAO\O_MAIOR_PROJETO_DE_TODOS_OS_TEMPOS\RELATORIOS_AUTO\img-assinatura.png"
    
    email.HTMLBody = f'''
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=iso-8859-1"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {a1}
    o\:* {a1}
    w\:* {a1}
    .shape {a1}
    </style><![endif]--><style><!--
    /* Font Definitions */
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a3}
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
    	{a4}
    span.EstiloDeEmail17
    	{a5}
    .MsoChpDefault
    	{a6}
    @page WordSection1
    	{a7}
    div.WordSection1
    	{a8}
    --></style><!--[if gte mso 9]><xml>
    <o:shapedefaults v:ext="edit" spidmax="1026" />
    </xml><![endif]--><!--[if gte mso 9]><xml>
    <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1" />
    </o:shapelayout></xml><![endif]-->
    </head>
    <body lang=PT-BR link="#0563C1" vlink="#954F72" style='word-wrap:break-word'>
    <div class=WordSection1>
    <p class=MsoNormal>
    <span style='font-size:12.0pt'>
    {nome}, {saudacao}!
    <br>
    <br>{corpo}{complemento_corpo_email}.
    <br>
    <br>Att.</span>
    <o:p>
    </o:p>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <span style='mso-fareast-language:PT-BR'>
    <o:p>&nbsp;</o:p>
    </span>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <b>
    <span style='font-size:12.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Relatórios Mis
    <br>
    </span>
    </b>
    <span style='font-size:10.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Envio automático 
    <br>
    <br>Bruno Abad Portugal Velasco -  Ramal 5288 / E-mail bruno.velasco@jarezende.com.br
    <br>Bruno de Almeida Frias Andriolli -  Ramal 5287 / E-mail bruno.andriolli@jarezende.com.br
    <br>Bruno Ezequiel Celidonio - Ramal 2734 / E-mail bruno.sousa@jarezende.com.br
    <br>João Vitor dos Santos Pinto - Ramal 2156 / E-mail joao.santos@jarezende.com.br    
    <br>
    <br>
    <img src="{img}" width="517" height="117">
    <br>
    <br>
    </span><i>
    <span style='font-size:8.0pt;line-height:106%;font-family:"Calibri Light",sans-serif;color:gray;mso-fareast-language:PT-BR'>
    Esta mensagem pode conter informação confidencial e/ou privilegiada. 
    Se você não for o destinatário ou a pessoa autorizada a receber esta mensagem, não pode usar, 
    copiar ou divulgar as informações nela contidas ou tomar qualquer ação baseada nessas informações. 
    Se você recebeu esta mensagem por engano, por favor, avise imediatamente o remetente, respondendo o e-mail, 
    e em seguida apague-o. 
    Agradecemos sua cooperação.
    <br>
    <br>
    </span>
    </i>
    <i>
    <span style='font-size:8.0pt;line-height:106%;color:gray;mso-fareast-language:PT-BR'>
    This message may contain confidential and/or privileged information. 
    If you are not the address or authorized to receive this for the address, you must not use, copy, 
    disclose or take any action base on this message or any information herein. 
    If you have received this message in error, please advise the sender immediately by reply e-mail and delete this message. 
    Thank you for your cooperation<o:p>
    </o:p></span></i></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>
    '''

    email.Send()    
    
    
    
def envia_email_validacao_sms_sicredi(qtd_tabela, qtd_ultimo_arquivo, qtd_atual, flag_cor, flag_palavra):
    email = outlook.CreateItem(0)
    email.To = 'planejamento-mis@jarezende.com.br'
    # email.CC = ''
    email.Subject = 'VALIDACAO SMS SICREDI {}'.format(data)
    a1 = '{behavior:url(#default#VML);}'
    a2 = '{font-family:"Cambria Math";ose-1:2 4 5 3 5 4 6 3 2 4;}'
    a3 = '{font-family:Lato;}'
    a4 = '{margin:0cm;t-size:11.0pt;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a5 = '{mso-style-type:personal-compose;t-family:"Calibri",sans-serif;or:windowtext;}'
    a6 = '{mso-style-type:export-only;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a7 = '{size:612.0pt 792.0pt;gin:70.85pt 3.0cm 70.85pt 3.0cm;}'
    a8 = '{page:WordSection1;}'
    img = r"\\192.168.200.14\JARezende\Administrativo\MIS\Planejamento_MIS\MIS\00-PESSOAL\JOAO\O_MAIOR_PROJETO_DE_TODOS_OS_TEMPOS\RELATORIOS_AUTO\img-assinatura.png"

    email.HTMLBody = f'''
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=iso-8859-1"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {a1}
    o\:* {a1}
    w\:* {a1}
    .shape {a1}
    </style><![endif]--><style><!--
    /* Font Definitions */
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a3}
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
    	{a4}
    span.EstiloDeEmail17
    	{a5}
    .MsoChpDefault
    	{a6}
    @page WordSection1
    	{a7}
    div.WordSection1
    	{a8}
    --></style><!--[if gte mso 9]><xml>
    <o:shapedefaults v:ext="edit" spidmax="1026" />
    </xml><![endif]--><!--[if gte mso 9]><xml>
    <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1" />
    </o:shapelayout></xml><![endif]-->
    </head>
    <body lang=PT-BR link="#0563C1" vlink="#954F72" style='word-wrap:break-word'>
    <div class=WordSection1>
    <p class=MsoNormal>
    <span style='font-size:12.0pt'>
    Prezados, {saudacao}!
    <br>
    <br>Segue resultado da validação do SMS SICREDI.
    <br>
    <br>QUANTIDADE DE REGISTROS NA TABELA: {qtd_tabela}
    <br>QUANTIDADE DE REGISTROS NO ÚLTIMO ARQUIVO: {qtd_ultimo_arquivo}
    <br>QUANTIDADE DE REGISTROS NO ARQUIVO ATUAL: {qtd_atual}
    <br>RESULTADO VALIDAÇÃO: <span style="color:{flag_cor}">{flag_palavra}</span>
    <br>
    <br>Att.</span>
    <o:p>
    </o:p>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <span style='mso-fareast-language:PT-BR'>
    <o:p>&nbsp;</o:p>
    </span>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <b>
    <span style='font-size:12.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Relatórios Mis
    <br>
    </span>
    </b>
    <span style='font-size:10.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Envio automático 
    <br>
    <br>Bruno Abad Portugal Velasco -  Ramal 5288 / E-mail bruno.velasco@jarezende.com.br
    <br>Bruno de Almeida Frias Andriolli -  Ramal 5287 / E-mail bruno.andriolli@jarezende.com.br
    <br>Bruno Ezequiel Celidonio - Ramal 2734 / E-mail bruno.sousa@jarezende.com.br
    <br>João Vitor dos Santos Pinto - Ramal 2156 / E-mail joao.santos@jarezende.com.br
    <br>
    <br>
    <img src="{img}" width="517" height="117">
    <br>
    <br>
    </span><i>
    <span style='font-size:8.0pt;line-height:106%;font-family:"Calibri Light",sans-serif;color:gray;mso-fareast-language:PT-BR'>
    Esta mensagem pode conter informação confidencial e/ou privilegiada. 
    Se você não for o destinatário ou a pessoa autorizada a receber esta mensagem, não pode usar, 
    copiar ou divulgar as informações nela contidas ou tomar qualquer ação baseada nessas informações. 
    Se você recebeu esta mensagem por engano, por favor, avise imediatamente o remetente, respondendo o e-mail, 
    e em seguida apague-o. 
    Agradecemos sua cooperação.
    <br>
    <br>
    </span>
    </i>
    <i>
    <span style='font-size:8.0pt;line-height:106%;color:gray;mso-fareast-language:PT-BR'>
    This message may contain confidential and/or privileged information. 
    If you are not the address or authorized to receive this for the address, you must not use, copy, 
    disclose or take any action base on this message or any information herein. 
    If you have received this message in error, please advise the sender immediately by reply e-mail and delete this message. 
    Thank you for your cooperation<o:p>
    </o:p></span></i></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>
    '''

    email.Send()      
    
    
        