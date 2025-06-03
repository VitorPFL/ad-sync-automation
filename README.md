# 🛠️ Automação de Gestão de Usuários no Active Directory Local

Este projeto foi desenvolvido com o objetivo de **organizar, padronizar e automatizar** a gestão de usuários no Active Directory (AD) local de uma empresa. A automação surgiu de uma necessidade real da equipe de TI, que enfrentava sérias dificuldades para manter controle sobre contas, acessos e dados dos colaboradores.

---

## 🎯 Problemas resolvidos

Antes desta solução, a empresa enfrentava:

- Contas de colaboradores demitidos ainda ativas no domínio
- Licenças de Microsoft 365 atribuídas a ex-funcionários, impedindo novos acessos
- Usuários criados em OUs erradas, dificultando localização e controle
- Informações incompletas ou inconsistentes nos perfis dos usuários no AD
- Falta de padronização no processo de criação e atualização de usuários

---

## ✅ Funcionalidades Atuais

- 📥 **Extração estruturada** de todos os usuários ativos do AD para planilha Excel
- 🔒 **Desativação automática** de contas de usuários desligados
- 🧾 **Atualização de atributos** como cargo, setor, empresa, nome, RE, CPF, e-mail, etc.
- 📁 **Movimentação de usuários** entre unidades organizacionais (OU) conforme a estrutura correta
- 🔄 **Integração com base de dados corporativa (TOTVS)** para garantir fidelidade das informações
- 🔐 **Uso de variáveis de ambiente** para proteger credenciais sensíveis

---

## 🌱 Melhorias Futuras

- 🆕 Criação automática de usuários a partir de chamados (integração com API de chamados)
- 🤝 Integração 100% automática: TOTVS → Chamado → Criação no AD
- 📦 Consolidação dos scripts em uma única aplicação com interface (CLI ou Web)
- 📊 Geração de logs e dashboards para rastreamento de alterações no AD

---

## 🧠 Tecnologias Utilizadas

- Python 3.x
- `ldap3` – comunicação com o AD via protocolo LDAP
- `pandas` – manipulação de planilhas
- `python-dotenv` – gerenciamento seguro de variáveis sensíveis
- Active Directory Local
- Microsoft Excel

---

## 🚀 Como Usar

1. Instale as dependências:
   ```bash
   pip install -r requirements.txt

---

## 👨‍💻 Autor

Vitor Pires Ferreira Leite
Pleno de Dados | Especialista em Automação de Processos, Integração de Sistemas e Soluções Corporativas

💡 Experiência com:

Python, VBA, SQL
Power BI (criação de dashboards interativos)
SharePoint, Power Apps, APIs REST
Integração entre plataformas corporativas (Metadados,TOTVS, AD, Microsoft 365)
Azure, Databricks

Meu Linkedin:
[LinkedIn](https://www.linkedin.com/in/vitor-ferreira-leite/)
