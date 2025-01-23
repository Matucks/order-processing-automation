# Order Processing Automation

Este repositório contém um script em Python projetado para automatizar o processamento de pedidos a partir de arquivos Excel. O script filtra os pedidos, separa-os por status (faturados e em progresso), realiza a formatação dos dados e salva os relatórios em tabelas Excel organizadas.

---

## Funcionalidades

- **Processamento de Arquivos Excel:**
  - Localiza e processa automaticamente o primeiro arquivo Excel encontrado na pasta de entrada.
- **Filtragem por Status:**
  - Gera relatórios de pedidos faturados.
  - Gera relatórios de pedidos em progresso.
- **Formatação de Tabelas:**
  - Salva os relatórios no formato de tabela Excel dinâmica.
- **Backup de Dados:**
  - Realiza um backup dos relatórios de progresso em uma pasta separada para fins de histórico.

---

## Requisitos

- **Python 3.8 ou superior.**
- **Bibliotecas necessárias:**
  - pandas
  - openpyxl
  - xlsxwriter

Para instalar as dependências, execute:
```bash
pip install pandas openpyxl xlsxwriter
```

---

## Configuração

1. Certifique-se de que os arquivos Excel a serem processados estejam localizados no diretório configurado como entrada:
   ```python
   INPUT_DIR = r"C:\\ProjectAutomation\\Input"
   ```

2. Configure os diretórios de saída e backup conforme necessário:
   ```python
   OUTPUT_DIR_BILLED = r"C:\\ProjectAutomation\\Output\\Billed"
   OUTPUT_DIR_PROGRESS = r"C:\\ProjectAutomation\\Output\\Progress"
   BACKUP_DIR = r"C:\\ProjectAutomation\\Backup"
   ```

3. Ajuste os filtros de pedidos para exclusão, se necessário:
   ```python
   orders_to_exclude = ["ORDER1", "ORDER2"]
   ```

---

## Como Executar

1. Clone o repositório para sua máquina local:
   ```bash
   git clone https://github.com/seu-usuario/order-processing-automation.git
   cd order-processing-automation
   ```

2. Execute o script principal:
   ```bash
   python main.py
   ```

3. Os relatórios serão gerados automaticamente nos diretórios configurados:
   - **Pedidos Faturados:** Salvo em `OUTPUT_DIR_BILLED`.
   - **Pedidos em Progresso:** Salvo em `OUTPUT_DIR_PROGRESS` e copiado para `BACKUP_DIR`.

---

## Estrutura dos Relatórios

### Pedidos Faturados
Inclui colunas como:
- **Order Number**
- **Event Date**
- **Branch Code**

Campos adicionais:
- **Destination Unit**, **Seller** e **Client** (em branco para preenchimento manual).

### Pedidos em Progresso
Inclui colunas como:
- **FAMILY**
- **MODEL**
- **OPTIONS**
- **COLOR**
- **BRANCH**
- **Week**
- **ORDER**
- **STATUS**

Campos adicionais:
- **Reservation Date**, **Destination Unit**, **Seller** e **Client** (em branco para preenchimento manual).

Ambos os relatórios são salvos no formato de tabela Excel dinâmica para facilitar a organização e a análise.

---

## Observações Importantes

- Certifique-se de que o nome das colunas nos arquivos de entrada esteja de acordo com o esperado pelo script.
- O script ignora arquivos temporários ou corrompidos que começam com `~$`.
- Ajuste as colunas removidas ou mapeadas conforme as regras de negócio.

---

## Contribuições

Contribuições são bem-vindas! Para relatar problemas, sugerir melhorias ou enviar pull requests, utilize a aba [Issues](https://github.com/seu-usuario/order-processing-automation/issues) no repositório.

---

## Licença

Este projeto está licenciado sob a MIT License.

---

## Autor

- **Gabriel Matuck**
- **E-mail:** [gabriel.matuck1@gmail.com](mailto:gabriel.matuck1@gmail.com)

---

Automatize o processamento de pedidos com agilidade e precisão!

