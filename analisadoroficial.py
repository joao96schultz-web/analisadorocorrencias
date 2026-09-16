#biblioteca
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
import smtplib
import os
from email.message import EmailMessage


#transformando em um diretorio
def criar_dataframe():

    df = pd.read_excel("relatorio.xlsx")

    print(df.head())

    #usando só as colunas necessárias
    dfres = df[["Alerta", "Investigação", "Equipamento", "Variável", "Data", "Hora", "Turno", "Planta"]]
    print(dfres.head())
    return dfres

def criar_vetores(dfres):

    #tabela ocorrência equipamento por turno
    tabelaequipa = pd.crosstab(index=[dfres["Equipamento"]],columns=dfres["Turno"],).fillna(0)
    tabelaequipa["Total"] = tabelaequipa.sum(axis=1)
    print(tabelaequipa)

    #investigação por turno

    tabelainvestiga = pd.crosstab(index=[dfres["Investigação"]],columns=dfres["Turno"],).fillna(0)
    tabelainvestiga.loc["Total"] = tabelainvestiga.sum(axis=0)
    print(tabelainvestiga)

    #variavel por máquina e turno
    tabelavalequipa = pd.crosstab(index=[dfres["Variável"], dfres["Equipamento"]], columns=dfres["Turno"]).fillna(0)
    tabelavalequipa["Total"] = tabelavalequipa.sum(axis=1)
    print(tabelavalequipa)

    #investigação por equipamento e turno
    tabelainvturno = pd.crosstab(index=[dfres["Investigação"], dfres["Equipamento"]], columns=dfres["Turno"]).fillna(0)
    tabelainvturno["Total"] = tabelainvturno.sum(axis=1)
    print(tabelainvturno)
    #variaveis
    investigacao = dfres["Investigação"].dropna().unique().tolist()
    dfres["Data"] = pd.to_datetime(dfres["Data"], dayfirst=True, errors="coerce")
    data = dfres["Data"].dropna()
    return (tabelaequipa, tabelainvestiga, tabelavalequipa, tabelainvturno, investigacao, data)
   

#funções para as tabelas
def criar_tabela(df):

    tabela_df = df.reset_index()
    estilo = ParagraphStyle("celula", fontSize=7, leading=8, alignment=TA_CENTER)

    # transforma todas as células em Paragraph
    dados = []

    # cabeçalho
    dados.append([Paragraph(str(col), estilo) for col in tabela_df.columns])

    # conteúdo
    for linha in tabela_df.values:
        dados.append([Paragraph(str(valor), estilo) for valor in linha])

    # largura disponível da página
    largura = A4[0] - 48 - 48

    # define larguras proporcionais
    n_colunas = len(tabela_df.columns)

    if n_colunas == 5:
        colWidths = [largura * 0.40, largura * 0.15, largura * 0.15, largura * 0.15, largura * 0.15]

    elif n_colunas == 6:
        colWidths = [largura * 0.44, largura * 0.26, largura * 0.075, largura * 0.075, largura * 0.075, largura * 0.075]

    else:
        colWidths = [largura / n_colunas] * n_colunas

    tabela = Table(dados, colWidths=colWidths, repeatRows=1)
    tabela.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("BACKGROUND", (0, 0), (-1, 0), colors.grey), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3), ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),]))

    return tabela

def desenhar_tabela(cnv, tabela, x, y, largura_pagina, margem_inferior=40):

    largura, altura = tabela.wrap(0, 0)

    # divide a tabela quando necessário
    partes = tabela.split(largura_pagina - 2 * x, y - margem_inferior)
    primeira = True

    for parte in partes:

        largura, altura = parte.wrap(0, 0)

        # se não couber na página atual
        if y - altura < margem_inferior:

            cnv.showPage()

            # nova página
            y = A4[1] - 40

        parte.drawOn(cnv, x, y - altura)

        y -= altura

        # espaço entre tabelas
        y -= 20

    return y

def criar_pdf(dfres, tabelaequipa, tabelainvestiga, tabelavalequipa, tabelainvturno, investigacao, data):
    #fazendo o pdf
    cnv = canvas.Canvas("Relatório_semanal.pdf", pagesize= A4)
    #título
    largura_pagina, altura_pagina = A4
    cnv.drawCentredString(largura_pagina / 2, 805.89, "Relatório Semanal")
    #data
    texto_data = f"Dia: {data.min().strftime('%d/%m/%Y')} até: {data.max().strftime('%d/%m/%Y')}"
    cnv.drawCentredString(largura_pagina / 2, 775, texto_data)
    #resumo das ocorrências
    maxocor = f"Número de ocorrências na semana: {len(dfres['Equipamento'])}"
    cnv.drawCentredString(largura_pagina / 2, 745, maxocor)
    maxinv = f"Número de investigações na semana: {len(investigacao)}"
    cnv.drawCentredString(largura_pagina / 2, 715, maxinv)
    porc = f"porcentagem de investigação: {((len(investigacao)/len(dfres['Equipamento']))*100):.2f} %"
    cnv.drawCentredString(largura_pagina / 2, 685, porc)
    ocoturnoa = dfres.loc[dfres["Turno"] == "A", "Alerta"].count()
    ocoturnob = dfres.loc[dfres["Turno"] == "B", "Alerta"].count()
    ocoturnoc = dfres.loc[dfres["Turno"] == "C", "Alerta"].count()
    invturnoa = dfres.loc[dfres["Turno"] == "A", "Investigação"].count()
    invturnob = dfres.loc[dfres["Turno"] == "B", "Investigação"].count()
    invturnoc = dfres.loc[dfres["Turno"] == "C", "Investigação"].count()
    cnv.drawCentredString(largura_pagina / 2, 655, f"Ocorrências no turno 1: {ocoturnoa}/ turno 2: {ocoturnob}/ turno 3: {ocoturnoc};")
    cnv.drawCentredString(largura_pagina / 2, 635, f"Investigações no turno 1: {invturnoa}/ turno 2: {invturnob}/ turno 3: {invturnoc};")
    cont_ocorrencias = dfres["Variável"].value_counts()
    maior_ocorrencia = cont_ocorrencias.index[0]
    quantidade = cont_ocorrencias.iloc[0]
    cnv.drawCentredString(largura_pagina / 2, 615, f"Alerta mais frequênte: {maior_ocorrencia}, com: {quantidade} vezes.")
    #criar tabelas para o pdf
    tabela_pdf1 = criar_tabela(tabelaequipa)

    tabela_pdf2 = criar_tabela(tabelainvestiga)

    tabela_pdf3 = criar_tabela(tabelavalequipa)

    tabela_pdf4 = criar_tabela(tabelainvturno)

    #anexos tabelas
    y = 595

    y = desenhar_tabela(cnv, tabela_pdf1, 48, y, largura_pagina)

    y = desenhar_tabela(cnv, tabela_pdf2, 48, y, largura_pagina)

    y = desenhar_tabela(cnv, tabela_pdf3, 48, y, largura_pagina)

    y = desenhar_tabela(cnv, tabela_pdf4, 48, y, largura_pagina)

    cnv.save()

def enviar_email_com_anexo():
  # Cria a mensagem
  msg = EmailMessage()
  msg['Subject'] = "Relatório Semanal"
  msg['From'] = "seuemail@gmail.com"
  msg['To'] = "nossoemail@hotmail.com"

  # Corpo do e-mail em texto puro (fallback) e HTML
  corpo_html = """
    <p>Bom dia,</p>
    <p>Segue relatório das occorências geradas no sistema durante a semana</p>
    """
  # Define o conteúdo HTML
  msg.set_content(
      'Este é um e-mail em formato HTML. Por favor, ative a exibição de HTML no'
      ' seu leitor de e-mail.'
  )
  msg.add_alternative(corpo_html, subtype='html')

  # --- ADICIONANDO O ANEXO ---
  caminho_anexo = 'F:/analisador/Relatório_semanal.pdf'  # Altere para o arquivo que deseja enviar

  if os.path.exists(caminho_anexo):
    # Obtém o nome do arquivo (ex: "seu_arquivo.pdf")
    nome_arquivo = os.path.basename(caminho_anexo)

    # Lê o arquivo em modo binário ('rb') e anexa
    with open(caminho_anexo, 'rb') as f:
      conteudo_arquivo = f.read()

      msg.add_attachment(
          conteudo_arquivo,
          maintype='application',  # Tipo genérico de arquivo
          subtype='octet-stream',  # Permite enviar qualquer extensão (pdf, xlsx, docx, png, etc.)
          filename=nome_arquivo,
      )
  else:
    print(f'Aviso: O arquivo {caminho_anexo} não foi encontrado.')

  # --- ENVIO VIA SMTP ---
  password = 'abcd efgh ijkl mnop'  # Senha de aplicativo do Gmail

  try:
    with smtplib.SMTP('smtp.gmail.com', 587) as s:
      s.starttls()
      s.login(msg['From'], password)
      s.send_message(msg)
      print('E-mail enviado com sucesso!')
  except Exception as e:
    print(f'Erro ao enviar e-mail: {e}')


dfres = criar_dataframe()
(tabelaequipa, tabelainvestiga, tabelavalequipa, tabelainvturno, investigacao, data,) = criar_vetores(dfres)
criar_pdf(dfres, tabelaequipa, tabelainvestiga, tabelavalequipa, tabelainvturno, investigacao, data)
enviar_email_com_anexo()
  
