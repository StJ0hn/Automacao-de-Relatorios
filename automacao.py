# ==============================================================================
# AUTOMAÇÃO DE RELATÓRIOS DE VENDAS
# ==============================================================================

# --- BIBLIOTECAS ---
# Ferramenta principal para manipulação de dados
import pandas as pd
# Ferramentas para construir e enviar e-mails
import smtplib
from email.message import EmailMessage

# --- CONFIGURAÇÃO DE E-MAIL (PREENCHER) ---
# ATENÇÃO: É necessário gerar uma "Senha de App" no Google e ativar a verificação em 2 etapas.
# Coloque o seu e-mail do Gmail que será usado para enviar os relatórios.
MEU_EMAIL = "emaildoremetente@gmail.com"
# Coloque a Senha de App de 16 letras gerada pelo Google.
MINHA_SENHA_APP = "senha_local_do_remetente"


# ==============================================================================
# FASE 1: AGREGAÇÃO DE DADOS (Organizar a Mesa de Trabalho)
# ==============================================================================
print("FASE 1: Iniciando a agregação de dados...")

# Carregando os arquivos de dados para a memória
tabela_vendas = pd.read_excel("Base de dados/Vendas.xlsx")
tabela_emails = pd.read_excel("Base de dados/Emails.xlsx")
tabela_lojas = pd.read_csv("Base de dados/Lojas.csv", sep=';', encoding='latin-1')

# Juntando todas as tabelas em uma única tabela mestra
# 1. Adiciona o nome da loja a cada venda, usando "ID Loja" como ponte
tabela_vendas = pd.merge(tabela_vendas, tabela_lojas, on="ID Loja")
# 2. Adiciona o gerente e e-mail a cada venda, usando "Loja" como ponte
tabela_vendas = pd.merge(tabela_vendas, tabela_emails, on="Loja")

print("-> Tabela mestra criada com sucesso.")


# ==============================================================================
# FASE 2: CÁLCULO DOS INDICADORES (Fazer as Contas)
# ==============================================================================
print("\nFASE 2: Calculando os indicadores anuais...")

# 1. Faturamento por Loja
faturamento_loja = tabela_vendas.groupby('Loja')[['Valor Final']].sum()

# 2. Quantidade de Vendas por Loja
quantidade_loja = tabela_vendas.groupby('Loja')[['Quantidade']].sum()

# 3. Ticket Médio por Loja
ticket_medio_loja = (faturamento_loja['Valor Final'] / quantidade_loja['Quantidade']).to_frame()
ticket_medio_loja = ticket_medio_loja.rename(columns={0: 'Ticket Médio'})

# Unindo os três indicadores em um painel de controle único
tabela_indicadores = faturamento_loja.merge(quantidade_loja, on='Loja').merge(ticket_medio_loja, on='Loja')
print("-> Painel de indicadores anuais concluído.")
print(tabela_indicadores)


# ==============================================================================
# FASE 3: RELATÓRIOS INDIVIDUAIS (Avisar os Gerentes)
# ==============================================================================
print("\nFASE 3: Iniciando envio de e-mails para os gerentes...")

# O laço de repetição que passa em cada loja
for loja in tabela_indicadores.index:
    # Extrai os indicadores da loja atual do painel
    faturamento = tabela_indicadores.loc[loja, 'Valor Final']
    quantidade = tabela_indicadores.loc[loja, 'Quantidade']
    ticket_medio = tabela_indicadores.loc[loja, 'Ticket Médio']

    # Pega o nome do gerente e o e-mail da tabela mestra original
    gerente = tabela_vendas.loc[tabela_vendas['Loja'] == loja, 'Gerente'].iloc[0]
    email_gerente = tabela_vendas.loc[tabela_vendas['Loja'] == loja, 'E-mail'].iloc[0]

    # --- Montagem do E-mail ---
    msg = EmailMessage()
    msg['Subject'] = f"OnePage Anual - Loja {loja}"
    msg['From'] = MEU_EMAIL
    msg['To'] = email_gerente

    corpo_email = f"""
    <p>Prezado(a) {gerente},</p>
    <p>Segue o relatório consolidado de desempenho da loja <strong>{loja}</strong>.</p>
    <ul>
        <li><strong>Faturamento Total:</strong> R${faturamento:,.2f}</li>
        <li><strong>Quantidade de Vendas:</strong> {quantidade:,}</li>
        <li><strong>Ticket Médio da Loja:</strong> R${ticket_medio:,.2f}</li>
    </ul>
    <p>Qualquer dúvida, estou à disposição.</p>
    <p>Att.,</p>
    <p>Robô de Automação</p>
    """
    msg.set_content(corpo_email, subtype='html')

    # --- Envio do E-mail ---
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(MEU_EMAIL, MINHA_SENHA_APP)
            smtp.send_message(msg)
            print(f"-> E-mail para a loja {loja} enviado com sucesso para {email_gerente}")
    except Exception as e:
        print(f"!! ERRO ao enviar e-mail para a loja {loja}: {e}")


# ==============================================================================
# FASE 4: RELATÓRIO GERAL (Apresentar o Resumo para a Diretoria)
# ==============================================================================
print("\nFASE 4: Gerando relatórios para a diretoria...")

# Encontrando o último dia de vendas na base de dados
dia_indicador = tabela_vendas['Data'].max()
vendas_dia = tabela_vendas[tabela_vendas['Data'] == dia_indicador]

# 1. Gerando o Ranking do Dia
faturamento_dia = vendas_dia.groupby('Loja')[['Valor Final']].sum()
ranking_dia = faturamento_dia.sort_values(by='Valor Final', ascending=False)
ranking_dia.to_excel("Ranking do Dia.xlsx")
print("-> Arquivo 'Ranking do Dia.xlsx' gerado.")

# 2. Gerando o Ranking Anual
ranking_anual = tabela_indicadores.sort_values(by='Valor Final', ascending=False)
ranking_anual.to_excel("Ranking Anual.xlsx")
print("-> Arquivo 'Ranking Anual.xlsx' gerado.")

# 3. Preparando o e-mail para a diretoria
melhor_loja_ano = ranking_anual.index[0]
pior_loja_ano = ranking_anual.index[-1]
faturamento_melhor_loja = ranking_anual.iloc[0, 0]
faturamento_pior_loja = ranking_anual.iloc[-1, 0]

# --- Montagem do E-mail da Diretoria ---
# O envio seria feito com a mesma lógica da Fase 3, mas com anexos.
# Por enquanto, apenas exibimos o corpo do e-mail.
corpo_email_diretoria = f"""
Prezados, bom dia.

O resultado consolidado do ano até o momento é:
Melhor Loja em Faturamento: {melhor_loja_ano} com R${faturamento_melhor_loja:,.2f}
Pior Loja em Faturamento: {pior_loja_ano} com R${faturamento_pior_loja:,.2f}

Seguem em anexo os rankings detalhados do ano e do último dia ({dia_indicador.strftime('%d/%m/%Y')}).

Qualquer dúvida, estou à disposição.

Att.,
Robô de Automação
"""

print("\n-> Preparando e-mail para a Diretoria:")
print(corpo_email_diretoria)

print("\n===============================")
print("AUTOMAÇÃO CONCLUÍDA COM SUCESSO!")
print("===============================")
