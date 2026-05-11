# Deploy do Dashboard de Operadores

## Arquivos necessários

O deploy precisa conter estes arquivos na raiz do projeto:

- `dashboard_operadores.py`
- `requirements.txt`
- `Cobmais-Eventos-908-2026050417.xlsx`
- `Pesquisa-Cliente-908-*.xlsx`
- `NOVA BASE RESULTADOS 2026.xlsm`

## Rodar localmente

```powershell
python -m streamlit run dashboard_operadores.py
```

## Deploy público pelo Streamlit Community Cloud

1. Crie um repositório no GitHub.
2. Envie os arquivos do projeto para esse repositório.
3. Acesse `https://share.streamlit.io`.
4. Conecte sua conta do GitHub.
5. Clique em `New app`.
6. Selecione o repositório, a branch e o arquivo principal:
   `dashboard_operadores.py`.
7. Clique em `Deploy`.

Ao final, o Streamlit gera um link público para acessar o dashboard.

## Atenção sobre dados

As planilhas contêm dados operacionais e podem conter dados sensíveis. Para um repositório público,
qualquer pessoa poderá baixar os arquivos. O recomendado é usar repositório privado ou uma base
anonimizada quando o link for compartilhado fora da empresa.
