# 📊 Automação de Relatórios de Vendas

Este projeto automatiza o processo de consolidação e envio de relatórios de desempenho de vendas para gerentes de loja e para a diretoria, utilizando Python e a biblioteca Pandas.

---

## 📄 Sobre o Projeto

O objetivo deste script é eliminar a tarefa manual e repetitiva de compilar dados de vendas de múltiplas fontes. A automação lê planilhas de vendas, lojas e e-mails, calcula indicadores-chave de desempenho (KPIs) para cada loja e distribui relatórios personalizados via e-mail.

---

## ✨ Funcionalidades

-   **Agregação de Dados**: Consolida informações de 3 arquivos distintos (`Vendas.xlsx`, `Lojas.csv`, `Emails.xlsx`) em uma única base de dados mestra.
-   **Cálculo de Indicadores**: Calcula automaticamente os seguintes KPIs para cada loja:
    -   Faturamento Total Anual
    -   Quantidade de Vendas Anual
    -   Ticket Médio Anual
-   **Envio de Relatórios por E-mail**: Dispara um e-mail individual e personalizado em HTML para cada gerente de loja, contendo os indicadores de desempenho específicos da sua loja.
-   **Geração de Rankings para a Diretoria**: Cria e salva duas planilhas Excel (`Ranking Anual.xlsx` e `Ranking do Dia.xlsx`) com o ranking das lojas por faturamento, prontas para análise pela diretoria.

---

## 🗃️ Estrutura dos Dados

O script espera encontrar os seguintes arquivos dentro da pasta `Base de dados/`:

-   **`Vendas.xlsx`**: Deve conter as colunas `ID Loja`, `Valor Final`, `Quantidade` e `Data`.
-   **`Lojas.csv`**: Deve conter as colunas `ID Loja` e `Loja`, separadas por ponto e vírgula (`;`).
-   **`Emails.xlsx`**: Deve conter as colunas `Loja`, `Gerente` e `E-mail`.

---

## 💻 Tecnologias Utilizadas

-   **Python 3**
-   **Pandas**: Para manipulação e análise dos dados.
-   **smtplib & email.message**: Bibliotecas padrão do Python para o envio dos e-mails.

---

## 🚀 Como Executar

Siga os passos abaixo para executar o projeto em sua máquina local.

1.  **Pré-requisitos**
    -   Certifique-se de ter o [Python 3](https://www.python.org/downloads/) instalado.

2.  **Clone o Repositório**
    ```bash
    git clone [https://URL-DO-SEU-REPOSITORIO.git](https://URL-DO-SEU-REPOSITORIO.git)
    cd nome-do-diretorio-do-projeto
    ```

3.  **Crie e Ative um Ambiente Virtual** (Recomendado)
    ```bash
    # Para Linux/macOS
    python3 -m venv .venv
    source .venv/bin/activate

    # Para Windows
    python -m venv .venv
    .\.venv\Scripts\activate
    ```

4.  **Instale as Dependências**
    ```bash
    pip install pandas openpyxl
    ```
    *`openpyxl` é necessário para que o Pandas possa ler e escrever arquivos `.xlsx`.*

5.  **Configure suas Credenciais**
    -   Abra o arquivo `automacao.py`.
    -   Preencha as variáveis `MEU_EMAIL` e `MINHA_SENHA_APP` com seu e-mail do Gmail e uma [Senha de App](https://support.google.com/accounts/answer/185833) gerada.
    -   **Atenção**: Não salve sua senha diretamente no código se for compartilhá-lo publicamente. Para projetos maiores, use variáveis de ambiente.

6.  **Execute o Script**
    ```bash
    python automacao.py
    ```

---

## 📈 Exemplo de Saída

Ao ser executado, o script irá:
-   Imprimir o status de cada fase da automação no terminal.
-   Enviar um e-mail para cada gerente de loja.
-   Gerar os arquivos `Ranking Anual.xlsx` e `Ranking do Dia.xlsx` no diretório principal.
