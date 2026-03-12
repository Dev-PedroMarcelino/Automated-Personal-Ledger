# AutoFinanças: Gerenciador Financeiro Pessoal Automatizado com Python & Excel

Este projeto é um bot de automação RPA (Robotic Process Automation) que simplifica o controle financeiro pessoal mensal. Ele elimina o trabalho repetitivo de criar e formatar planilhas de gastos, gerando automaticamente arquivos Excel estilizados, com fórmulas nativas e persistência de dados do mês anterior.

## 🚀 Funcionalidades

- **Automação de Fechamento Mensal:** Gera uma nova planilha Excel a cada mês via gatilho temporal do Sistema Operacional.
- **Entrada Dinâmica de Dados:** Utiliza uma interface visual (Tkinter) para solicitar o valor da fatura do cartão de crédito no início do ciclo.
- **Design Premium Autônomo:** Formata automaticamente o Excel com cabeçalhos estilizados, larguras automáticas, bordas e formato de moeda (R$) usando a biblioteca `openpyxl`.
- **Fórmulas Nativas do Excel:** Injeta fórmulas de soma (`=SUM()`) e saldo diretamente no arquivo, permitindo que os totais se atualizem automaticamente caso o usuário edite a planilha no Excel.
- **Persistência de Estado (ETL):** Ao criar uma nova planilha, o motor busca a planilha do mês anterior, copia as categorias e o planejamento, zerando apenas a coluna de gastos reais.
- **Arquitetura de Configuração:** Isolamento total dos dados pessoais do usuário em um arquivo `config.py` separado do motor lógico.

## 🛠️ Tecnologias Utilizadas

- **Python 3.x**
- **Pandas:** Para manipulação estruturada dos dados (ETL).
- **OpenPyxl:** Para automação avançada de formatação e injeção de fórmulas no Excel.
- **Plyer:** Para notificações nativas do Sistema Operacional.
- **Tkinter:** Para interfaces simples de entrada de dados.

---

## 💻 Passo a Passo para Iniciar a Automação na Sua Máquina (Windows)

Siga este guia detalhado para clonar e configurar o projeto no seu computador.

### Fase 1: Pré-requisitos

1.  **Instalar o Python:** Baixe e instale a versão mais recente do Python em [python.org](https://www.python.org/downloads/).
    * **ALERTA CRÍTICO:** Na primeira tela do instalador, você **DEVE** marcar a caixa **"Add python.exe to PATH"**. Se não marcar, os comandos no terminal não funcionarão.

### Fase 2: Instalação do Projeto e Dependências

1.  **Baixar o Projeto:** Clique no botão verde "Code" acima e selecione "Download ZIP". Extraia o arquivo em uma pasta segura no seu computador (ex: `C:\MeusProjetos\GerenciadorFinanceiro`).
2.  **Abrir o Terminal:** Pressione a tecla `Windows`, digite `cmd` e aperte `Enter`.
3.  **Navegar até a Pasta:** No terminal, use o comando `cd` para entrar na pasta onde extraiu o projeto.
    ```cmd
    cd C:\MeusProjetos\GerenciadorFinanceiro
    ```
4.  **Instalar Bibliotecas:** Digite o seguinte comando para instalar todas as dependências necessárias listadas no arquivo `requirements.txt`.
    ```cmd
    pip install -r requirements.txt
    ```

### Fase 3: Configuração dos Seus Dados

1.  Abra o arquivo `config.py` com qualquer editor de texto (Bloco de Notas, VS Code, etc.).
2.  Edite os valores das variáveis de acordo com a sua realidade pessoal:
    * `RECEITA_LIQUIDA`: Digite o valor do seu salário líquido (use ponto para decimais, ex: `2500.00`).
    * `MATRIZ_GASTOS_PADRAO`: Adicione, remova ou edite as linhas para refletir seus gastos fixos e metas de poupança (câmbio). *Mantenha a sintaxe da lista do Python.*

### Fase 4: Homologação (Teste Manual)

1.  No terminal (ainda dentro da pasta do projeto), force a execução manual do motor para verificar se tudo está configurado corretamente:
    ```cmd
    python motor_financeiro.py
    ```
2.  A janela solicitando o valor do cartão deve aparecer. Digite um valor de teste e clique em OK.
3.  Uma notificação do Windows deve aparecer. Verifique se a pasta `planilhas_geradas` (ou o nome que configurou) foi criada e se o arquivo Excel está lá, perfeitamente formatado e com os totais calculados.

### Fase 5: Agendamento no Sistema Operacional (Windows Task Scheduler)

Para que o motor rode automaticamente todo mês (ex: dia 29), delegamos o gatilho para o Windows.

1.  Mapeie dois caminhos importantes:
    * **Caminho do Python:** No terminal, digite `where python` e copie o caminho do executável.
    * **Caminho do Script:** É o caminho completo do seu arquivo `motor_financeiro.py`.
2.  Pressione `Windows`, busque por **Agendador de Tarefas** e abra-o.
3.  Clique em **Criar Tarefa Básica...** no painel direito.
4.  **Nome:** `Automação Caixa Mensal`. Avance.
5.  **Disparador:** Escolha **Mensalmente**. Avance.
    * Marque "Selecionar todos os meses" e defina o dia para `29` (ou o dia que preferir). Avance.
6.  **Ação:** Deixe **Iniciar um programa**. Avance.
7.  **Iniciação:**
    * No campo **Programa/script**, cole o caminho completo do Python que você mapeou (ex: `C:\Users\...\python.exe`).
    * No campo **Adicione argumentos (opcional)**, cole o caminho completo do seu script `motor_financeiro.py`. Avance.
8.  Revise e clique em **Concluir**.

---
**Status:** Automação em produção. O sistema agora gerencia o controle de danos financeiros mensalmente sem esforço manual.