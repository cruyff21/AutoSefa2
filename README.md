Automação para Contestação de Documentos no Site da SEFA
Este projeto automatiza o processo de envio de documentos no site da Secretaria da Fazenda (SEFA) usando Python. Ele realiza login, navega até a inscrição estadual, seleciona os documentos a serem enviados e preenche automaticamente os dados de acordo com uma planilha do Excel.

🚀 Funcionalidades

Leitura de Dados: Carrega os dados necessários de um arquivo Excel, como CPF, senha, inscrição estadual, e período de envio.
Automação Web: Utiliza Selenium para interagir com o site da SEFA e realizar as ações automaticamente.
Atualização do Status: Registra na planilha o status do envio para cada inscrição estadual.


🛠️ Tecnologias Utilizadas

Python: Linguagem principal do projeto.
Selenium: Automação do navegador.
openpyxl: Manipulação de arquivos Excel.
WebDriver Manager: Gerenciamento do driver do navegador.
Google Chrome: Navegador utilizado na automação.


📋 Pré-requisitos

Python 3.9+ instalado na máquina.
Navegador Google Chrome e o respectivo ChromeDriver.
Arquivo Excel (Empresas_sefa.xlsx) no formato esperado.
Instalação das Dependências
Use o comando abaixo para instalar as bibliotecas necessárias:

bash
Copiar código
pip install selenium openpyxl webdriver-manager
🗂️ Estrutura do Projeto
bash
Copiar código
📂 projeto_sefa
├── Empresas_sefa.xlsx       # Planilha com os dados das empresas
├── main.py                  # Script principal do projeto
├── README.md                # Documentação do projeto
📋 Formato do Arquivo Excel
O arquivo Empresas_sefa.xlsx deve conter os seguintes campos a partir da linha 2:


A	B	C	D	E	F	G

Nº	CPF	Senha	Inscrição	Data Inicial	Data Final	Status

CPF: CPF do usuário para login.

Senha: Senha de acesso.

Inscrição: Inscrição estadual a ser selecionada.

Data Inicial e Data Final: Período para o envio de documentos.

Status: Atualizado automaticamente pelo script.


▶️ Como Executar
      Certifique-se de que o arquivo Empresas_sefa.xlsx esteja no mesmo diretório do script.
        Execute o script principal:
        bash
          Copiar código
          python main.py

A automação irá:
  Ler os dados do Excel.
    Realizar login no site da SEFA.
      Navegar para a inscrição estadual.
        Preencher e enviar os documentos conforme o período.
          Atualizar o status na planilha.


🖋️ Autor
Desenvolvido por Victor Estevam.
https://www.linkedin.com/in/victor-estevam-baia-002671237/ | adrieneestevam50@gmail.com

📝 Licença
Este projeto é licenciado sob a licença MIT. Consulte o arquivo LICENSE para mais detalhes.
