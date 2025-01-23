Order Processing Automation

Este repositório contém um script em Python projetado para automatizar o processamento de pedidos a partir de arquivos Excel. O script filtra os pedidos, separa-os por status (faturados e em progresso), realiza a formatação dos dados e salva os relatórios em tabelas Excel organizadas.

Funcionalidades

Processamento de arquivos Excel: Localiza e processa automaticamente o primeiro arquivo Excel encontrado na pasta de entrada.

Filtragem por status:

Cria relatórios de pedidos faturados.

Cria relatórios de pedidos em progresso.

Formatação de tabelas: Os relatórios são salvos no formato de tabela Excel dinâmica.

Backup de dados: Realiza um backup dos relatórios de progresso em uma pasta separada para fins de histórico.

Requisitos

Python 3.8 ou superior.

Bibliotecas Python necessárias:

pandas

openpyxl

xlsxwriter

Para instalar as dependências, execute:

pip install pandas openpyxl xlsxwriter

Configuração

Certifique-se de que os arquivos Excel a serem processados estejam localizados no diretório configurado como entrada:

INPUT_DIR = r"C:\\ProjectAutomation\\Input"

Configure os diretórios de saída e backup conforme necessário:

OUTPUT_DIR_BILLED = r"C:\\ProjectAutomation\\Output\\Billed"
OUTPUT_DIR_PROGRESS = r"C:\\ProjectAutomation\\Output\\Progress"
BACKUP_DIR = r"C:\\ProjectAutomation\\Backup"

Ajuste os filtros de pedidos para exclusão, se necessário:

orders_to_exclude = ["ORDER1", "ORDER2"]

Como Executar

Clone o repositório para sua máquina local:

git clone https://github.com/seu-usuario/order-processing-automation.git
cd order-processing-automation

Execute o script principal:

python main.py

Os relatórios serão gerados automaticamente nos diretórios configurados:

Pedidos faturados: Salvo em OUTPUT_DIR_BILLED.

Pedidos em progresso: Salvo em OUTPUT_DIR_PROGRESS e copiado para BACKUP_DIR.

Estrutura dos Relatórios

Pedidos Faturados: Inclui colunas como:

Order Number, Event Date, Branch Code.

Campos adicionais como Destination Unit, Seller e Client (em branco para preenchimento manual).

Pedidos em Progresso: Inclui colunas como:

FAMILY, MODEL, OPTIONS, COLOR, BRANCH, Week, ORDER, STATUS.

Campos adicionais como Reservation Date, Destination Unit, Seller e Client (em branco para preenchimento manual).

Ambos os relatórios são salvos no formato de tabela Excel dinâmica para facilitar a organização e a análise.

Observações Importantes

Certifique-se de que o nome das colunas nos arquivos de entrada esteja de acordo com o esperado pelo script.

O script ignora arquivos temporários ou corrompidos que começam com ~$.

Ajuste as colunas removidas ou mapeadas conforme as regras de negócio.

Contribuições

Contribuições são bem-vindas! Para relatar problemas, sugerir melhorias ou enviar pull requests, utilize a aba "Issues" no repositório.

Licença

Este projeto está licenciado sob a MIT License.

Autor: Gabriel Matuck

Contato: gabriel.matuck1@gmail.com
