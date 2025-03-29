import pandas as pd
import smtplib
import os
from email.message import EmailMessage
from plyer import notification

def analisar_vendas(arquivo_csv):
    # Carregar o arquivo CSV
    df = pd.read_csv(arquivo_csv)
    
    # Verificar se o arquivo possui as colunas esperadas
    if not {'Produto', 'Quantidade', 'Total'}.issubset(df.columns):
        raise ValueError("O arquivo CSV deve conter as colunas: Produto, Quantidade e Total")
    
    # Agrupar por produto
    resumo_vendas = df.groupby('Produto').agg({'Quantidade': 'sum', 'Total': 'sum'}).reset_index()
    
    # Produto mais vendido
    produto_mais_vendido = resumo_vendas.loc[resumo_vendas['Quantidade'].idxmax(), 'Produto']
    
    # Total geral de vendas
    total_vendas = df['Total'].sum()
    
    return resumo_vendas, produto_mais_vendido, total_vendas

def salvar_relatorio(resumo_vendas, arquivo_excel):
    with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
        resumo_vendas.to_excel(writer, sheet_name='Resumo de Vendas', index=False)

def enviar_notificacao(produto, total):
    mensagem = f"Produto mais vendido: {produto}\nTotal de vendas: R$ {total:,.2f}"
    notification.notify(
        title="Resumo de Vendas",
        message=mensagem,
        timeout=10
    )

def enviar_email(arquivo_excel, destinatario, remetente, senha):
    msg = EmailMessage()
    msg['Subject'] = "Relatório de Vendas"
    msg['From'] = remetente
    msg['To'] = destinatario
    msg.set_content("Segue em anexo o relatório de vendas.")
    
    with open(arquivo_excel, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(arquivo_excel)
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
    
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(remetente, senha)
        server.send_message(msg)
    
    print("E-mail enviado com sucesso!")

if __name__ == "__main__":
    try:
        arquivo_csv = "vendas.csv"  # Substitua pelo nome correto do arquivo
        arquivo_excel = "relatorio.xlsx"
        destinatario = "destinatario@example.com"  # Substitua pelo e-mail do destinatário
        remetente = "seuemail@gmail.com"  # Substitua pelo seu e-mail
        senha = "sua_senha"  # Substitua pela senha do seu e-mail (use autenticação segura, como app password)
        
        resumo, produto, total = analisar_vendas(arquivo_csv)
        salvar_relatorio(resumo, arquivo_excel)
        enviar_notificacao(produto, total)
        enviar_email(arquivo_excel, destinatario, remetente, senha)
    except Exception as e:
        print(f"Erro: {e}")


