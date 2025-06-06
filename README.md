<h1 align="center">Bot de Registro de Horários no Discord</h1>

Este projeto é um bot do Discord que registra horários de entrada e saída dos usuários, armazenando os dados em um banco SQLite e gerando relatórios em Excel.

## ✨ Funcionalidades

- `!entrada` → Registra a entrada do usuário no banco.
- `!saida <resumo>` → Registra a saída do usuário com um resumo.
- `!criar_resumo` → Gera um arquivo Excel (.xlsx) com todos os registros salvos.

## 📦 Requisitos

- Python 3.x
- Dependências Python:
  - discord.py
  - pandas
  - openpyxl
  - python-dotenv

## ⚙️ Como rodar

1. Clone o repositório:
   ```
   git clone <URL-do-repo>
   ```

2. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```

3. Crie um arquivo `.env` na raiz do projeto e adicione:
   ```
   DISCORD_TOKEN=seu_token_aqui
   ```

4. Execute o bot:
   ```
   python bot.py
   ```

## 👨‍💻 Feito por [Gabriel Ribeiro](https://www.linkedin.com/in/fulldevgabriel/)


