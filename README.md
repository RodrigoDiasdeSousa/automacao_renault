Este projeto consiste em um script Python desenvolvido para automatizar a análise de indicadores de qualidade (GMB), gerar evidências screenshots e realizar o disparo de e-mails personalizados via Microsoft Outlook.

Acompanhamento dispersão.xlsx: A planilha principal contendo os dados e notas das concessionárias.

Emails grupo.xlsm: A base de dados com os destinatários (Para/Cc) de cada grupo.

A imagem do logo da Renault/Assinatura que vai no rodapé do e-mail.

O script abre a planilha de dados (Acompanhamento dispersão.xlsx) usando a biblioteca openpyxl (para leitura de dados) e também via win32com (para tirar fotos).

Seleciona a coluna com a data mais recente com biblioteca datetime

Seleciona as concessionárias com <780 pontos 

Abre o excel em segundo plano e copia a imagem do intervalo selecionado

Insere a imagem no corpo do email junto com uma mensagem em HTML e anexa a assinatura junto

O código está configurado com .Display(), ou seja, ele cria os e-mails e deixa na tela para revisão. Para enviar automático, altera-se para .Send().
