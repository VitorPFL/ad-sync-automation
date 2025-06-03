# ad-sync-automation
Scripts de automação para sincronização de dados entre o Active Directory local, Microsoft 365 e TOTVS. O objetivo é manter as informações de usuários atualizadas, desativar contas inativas e garantir organização no ambiente corporativo.


# AD Sync Automation

Este repositório contém scripts desenvolvidos para automatizar a sincronização e organização dos dados de usuários entre o Active Directory local, Microsoft 365 e a base de colaboradores do TOTVS.

---

## Objetivo

Devido à desorganização de usuários no ambiente corporativo — contas ativas mesmo após demissão, usuários duplicados entre o AD e o 365, e uso desnecessário de licenças — foi criado um fluxo automatizado que:

- Extrai os dados atualizados da base de colaboradores (TOTVS)
- Compara com os dados do AD local
- Atualiza informações como função, local, gerente, data de admissão, etc.
- Desativa automaticamente contas de colaboradores desligados
- Mantém o ambiente sincronizado com o mínimo de intervenção manual

---

## Tecnologias Utilizadas

- Python 3.x  
- `ldap3`  
- `pandas`  
- Active Directory (via LDAP)
- Microsoft 365
- TOTVS (como origem das informações de colaboradores)

---

## Estrutura dos Arquivos

| Arquivo                       | Função                                                              |
|-------------------------------|-----------------------------------------------------------------------|
| `extrair_dados_ad.py`         | Extrai informações completas do AD Local                              |
| `desativar_usuarios.py`       | Desativa automaticamente contas de usuários desligados                |
| `atualizar_usuarios_ad.py`    | Atualiza informações de usuários ativos com base na planilha do TOTVS |

---

## Como Usar

1. Configure os parâmetros de acesso no início dos scripts (`servidor`, `usuário`, `senha`, `base_dn`)
2. Execute `extrair_dados_ad.py` para gerar a base atual de usuários
3. Utilize a base do TOTVS como referência para validar e alimentar as próximas ações
4. Execute `desativar_usuarios.py` para desabilitar usuários desligados
5. Execute `atualizar_usuarios_ad.py` para atualizar as informações dos usuários ativos

---

##  Autor

**Vitor Pires Ferreira Leite**  
Pleno de Dados | Especialista em automação e integração corporativa  
[LinkedIn](https://www.linkedin.com/in/vitor-ferreira-leite)

---

## Melhorias Futuras

- Integração direta com a API Graph para sincronizar com o Azure AD
- Interface visual para controle manual das ações
- Logs e dashboards de auditoria para alterações no AD
- Juntar todos os códigos em um código unico
- Não precisar de validação, deixar com menos contato humano.

