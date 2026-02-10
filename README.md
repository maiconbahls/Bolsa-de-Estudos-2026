# 🎓 Sistema de Gestão de Bolsas de Estudos - COCAL

Sistema completo para gerenciamento de bolsas de estudos com dashboard interativo, conferência mensal de pagamentos e análise de dados.

## 📋 Funcionalidades

- ✅ **Dashboard Estratégico**: Visualização de KPIs e métricas importantes
- ✅ **Gestão de Bolsistas**: Cadastro, edição e acompanhamento completo
- ✅ **Conferência Mensal**: Sistema de aprovação de pagamentos
- ✅ **Histórico de Pagamentos**: Análise temporal e por diretoria
- ✅ **Integração com Organograma**: Mapeamento automático de diretorias via código local
- ✅ **Relatórios Exportáveis**: Download em Excel de todos os dados

## 🚀 Deploy no Streamlit Community Cloud

### Pré-requisitos
1. Conta no [GitHub](https://github.com)
2. Conta no [Streamlit Community Cloud](https://streamlit.io/cloud)

### Passo a Passo

#### 1. Preparar o Repositório GitHub

```bash
# Inicializar repositório Git (se ainda não foi feito)
git init

# Adicionar todos os arquivos
git add .

# Fazer o primeiro commit
git commit -m "Initial commit - Sistema de Bolsas COCAL"

# Criar repositório no GitHub e conectar
git remote add origin https://github.com/SEU_USUARIO/NOME_DO_REPOSITORIO.git
git branch -M main
git push -u origin main
```

#### 2. Deploy no Streamlit Cloud

1. Acesse [share.streamlit.io](https://share.streamlit.io)
2. Faça login com sua conta GitHub
3. Clique em **"New app"**
4. Selecione:
   - **Repository**: Seu repositório
   - **Branch**: main
   - **Main file path**: app.py
5. Clique em **"Deploy!"**

#### 3. Configurar Secrets (Dados Sensíveis)

Se você tiver credenciais ou dados sensíveis, configure em:
- Settings → Secrets
- Adicione no formato TOML:

```toml
# Exemplo de secrets
[database]
connection_string = "sua_string_de_conexao"

[google_sheets]
credentials = '''
{
  "type": "service_account",
  ...
}
'''
```

## 📦 Estrutura do Projeto

```
SISTEMA_BOLSAS_DEESTUDOS/
├── app.py                      # Aplicação principal
├── requirements.txt            # Dependências Python
├── .streamlit/
│   └── config.toml            # Configurações do Streamlit
├── .gitignore                 # Arquivos ignorados pelo Git
├── BASES.BOLSAS/              # Arquivos de dados
│   ├── BASE.BOLSAS.2025.xlsx
│   ├── BASE.PAGAMENTOS.xlsx
│   └── ORGANOGRAMA.xlsx
├── static/
│   └── style.css              # Estilos customizados
└── backups/                   # Backups automáticos do banco

```

## 🔧 Executar Localmente

```bash
# Instalar dependências
pip install -r requirements.txt

# Executar aplicação
streamlit run app.py
```

A aplicação estará disponível em `http://localhost:8501`

## ⚙️ Tecnologias Utilizadas

- **Python 3.11+**
- **Streamlit**: Framework web
- **Pandas**: Manipulação de dados
- **Plotly**: Gráficos interativos
- **SQLite**: Banco de dados local
- **openpyxl**: Leitura/escrita de Excel
- **streamlit-aggrid**: Tabelas interativas avançadas

## 📊 Banco de Dados

O sistema utiliza SQLite com as seguintes tabelas:

- `bolsistas`: Cadastro de bolsistas
- `pagamentos`: Controle mensal de pagamentos
- `historico_pagamentos`: Histórico importado de Excel
- `observacoes`: Anotações e documentos anexados
- `orcamento`: Metas orçamentárias por diretoria

## 🔐 Segurança

- ✅ Backups automáticos antes de operações críticas
- ✅ Validação de dados em todas as entradas
- ✅ Proteção contra SQL Injection
- ✅ Controle de acesso por sessão

## 📝 Notas Importantes

### Arquivos de Dados
Os arquivos da pasta `BASES.BOLSAS/` **NÃO** são enviados para o GitHub por questões de segurança (estão no `.gitignore`). 

Para deploy em produção, você precisará:
1. Fazer upload manual dos arquivos Excel, OU
2. Configurar integração com Google Sheets, OU
3. Usar um banco de dados em nuvem (PostgreSQL, MySQL, etc.)

### Banco de Dados em Produção
Para produção, recomenda-se migrar de SQLite para um banco de dados mais robusto como PostgreSQL.

## 🆘 Suporte

Para dúvidas ou problemas:
1. Verifique os logs do Streamlit Cloud
2. Revise as configurações de secrets
3. Confirme que todos os arquivos necessários estão no repositório

## 📄 Licença

Uso interno - COCAL

---

**Desenvolvido com ❤️ para COCAL**
