# 📊 Tratador de Extrato Genial - Web App

Aplicação web para processar extratos financeiros de forma simples e visual.

## 🚀 Como Rodar Localmente

### 1. Instalar as dependências
```bash
pip install -r requirements.txt
```

### 2. Rodar o app
```bash
streamlit run app.py
```

O app abrirá automaticamente no seu navegador em `http://localhost:8501`

---

## 📋 Como Usar

1. **Faça upload** do arquivo Excel (.xlsx) com o extrato
2. **Clique** em "⚡ PROCESSAR ARQUIVO"
3. **Visualize** os dados e estatísticas na tela
4. **Baixe** o arquivo Excel tratado automaticamente

---

## 🌐 Compartilhar com a Equipe (Grátis)

### Opção 1: Streamlit Cloud (Recomendado)

1. Crie uma conta em [GitHub](https://github.com)
2. Suba este repositório para o GitHub
3. Acesse [share.streamlit.io](https://share.streamlit.io)
4. Conecte seu GitHub e selecione o repositório
5. Pronto! Sua equipe terá um link para acessar

### Opção 2: Rede Local

Para uso apenas na sua rede interna:

```bash
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

Sua equipe acessará pelo seu IP: `http://SEU_IP:8501`

---

## 📁 Estrutura dos Arquivos

```
├── app.py                  # Aplicação web principal
├── requirements.txt        # Dependências do projeto
├── README.md               # Este arquivo
└── script copy.py          # Script original (referência)
```

---

## ✅ Formato Esperado do Excel de Entrada

O arquivo Excel deve conter as seguintes colunas:

| Coluna | Descrição |
|--------|-----------|
| `Data` | Data da transação |
| `HISTORICO` | Tipo de transação |
| `Valor` | Valor financeiro |
| `HISTORICO DE LANÇAMENTO` | Descrição do lançamento |

---

## 🔧 O que o App Faz

- ✅ Agrupa PAY IN e PAY OUT por dia (soma os valores)
- ✅ Mantém outras transações individualmente
- ✅ Formata o Excel com cores e bordas
- ✅ Destaca visualmente entradas (verde) e saídas (vermelho)
- ✅ Congela o cabeçalho para facilitar navegação
