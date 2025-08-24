# AutoMail-Sales-Consolidator
Um script em Python que automatiza a consolidação de arquivos CSV de vendas em uma única planilha Excel e envia o relatório atualizado por e-mail automaticamente via Outlook.

### 📌 Descrição
Este projeto automatiza o processo de consolidação de arquivos CSV em uma única planilha Excel (`Vendas.xlsx`) e o envio automático do relatório por e-mail através do Outlook.  

Ideal para equipes que precisam de relatórios de vendas diários de forma rápida e prática.

---

### 🚀 Funcionalidades
- Lê automaticamente todos os arquivos CSV dentro da pasta `arquivos_csv`.
- Consolida todos os dados em uma única planilha Excel.
- Ordena as vendas pela data.
- Gera o arquivo `Vendas.xlsx` pronto para uso.
- Envia o relatório por e-mail usando o Outlook (com corpo e assunto automáticos).

---

### 🛠️ Pré-requisitos
- Python 3.9+
- Ter o **Outlook configurado e funcionando** no computador.
- Bibliotecas necessárias (instale rodando o comando abaixo no terminal):

```bash
pip install pandas numpy openpyxl pywin32
```

---

### 📂 Estrutura de Pastas

📦 Projeto<br>
 ┣ 📂 arquivos_csv   # Pasta onde ficam os arquivos CSV de entrada<br>
 ┣ 📜 codigo.py      # Script principal<br>
 ┣ 📜 Vendas.xlsx    # Gerado automaticamente após a execução<br>
 ┣ 📜 informacoes.txt # Descrição do projeto

---

### ▶️ Como Usar

Coloque todos os arquivos CSV dentro da pasta arquivos_csv.

Execute o script no terminal:

```bash
python codigo.py
```

O programa irá:

- Consolidar os arquivos.
- Gerar o Vendas.xlsx.
- Enviar o relatório por e-mail automaticamente.
