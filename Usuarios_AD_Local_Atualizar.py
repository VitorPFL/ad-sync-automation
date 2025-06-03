import ldap3
import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv(dotenv_path="AD.env")


# Configuração do servidor AD
SERVIDOR_AD = os.getenv("SERVIDOR_AD")
USUARIO_AD = os.getenv("USUARIO_AD")
SENHA_AD = os.getenv("SENHA_AD")
BASE_DN = os.getenv("BASE_DN")

# Caminho do arquivo Excel
caminho_arquivo = os.path.join(os.path.expanduser("~"), "Downloads", "Atualização AD.xlsx")

# Carregar a planilha
df = pd.read_excel(caminho_arquivo)

# Criar conexão com o AD
server = ldap3.Server(SERVIDOR_AD, get_info=ldap3.ALL)
conn = ldap3.Connection(server, user=USUARIO_AD, password=SENHA_AD, auto_bind=True)

if conn.bind():
    print("✅ Conexão com o Active Directory estabelecida com sucesso!\n")

    for index, row in df.iterrows():
        login_ad = str(row['Login']).strip()
        first_name = str(row['Primeiro Nome Dcolab']).strip()
        last_name = str(row['Sobrenome Dcolab']).strip()
        email = str(row['Email 365']).strip()
        user_principal_name = str(row['Email 365']).strip()
        title = str(row['Funcao']).strip()
        department = str(row['Secao']).strip()
        company = str(row['Empresa']).strip()
        user_logon_name_pre2000 = str(row["User Logon name pre"]).strip()
        display_name = str(row["Nome Dcolab"]).strip()
        CPF = str(row["CPF"]).strip()
        Supervisor = str(row["Supervisor"]).strip()
        RE = str(row["RE"]).strip()
        Descricao = str(row["Descrição"]).strip()


        # Buscar o usuário pelo sAMAccountName
        conn.search(
            BASE_DN,
            f"(&(objectClass=user)(sAMAccountName={login_ad})(!(objectCategory=computer)))",
            attributes=["cn", "sAMAccountName", "mail"]
        )

        if conn.entries:
            usuario = conn.entries[0]
            print(f"🔄 Atualizando {login_ad} - {usuario.cn.value}")

            atributos_para_atualizar = {
                #'givenName': [(ldap3.MODIFY_REPLACE, [first_name])],
                #'sn': [(ldap3.MODIFY_REPLACE, [last_name])],
                # 'mail': [(ldap3.MODIFY_REPLACE, [email])],
                # 'userPrincipalName': [(ldap3.MODIFY_REPLACE, [user_principal_name])],
                # 'sAMAccountName': [(ldap3.MODIFY_REPLACE, [user_logon_name_pre2000])],
                #'title': [(ldap3.MODIFY_REPLACE, [title])],
                #'department': [(ldap3.MODIFY_REPLACE, [department])],
                #'company': [(ldap3.MODIFY_REPLACE, [company])],
                #'displayName': [(ldap3.MODIFY_REPLACE, [display_name])],
                'employeeNumber': [(ldap3.MODIFY_REPLACE, [CPF])],
                #'manager': [(ldap3.MODIFY_REPLACE, [Supervisor])],
                'employeeID': [(ldap3.MODIFY_REPLACE, [RE])],
                'employeeType': [(ldap3.MODIFY_REPLACE, [Descricao])]
            }

            if conn.modify(usuario.entry_dn, atributos_para_atualizar):
                print(f"✅ Usuário {login_ad} atualizado com sucesso!\n")
                print("📌 Campos atualizados:")
                for campo, valor in atributos_para_atualizar.items():
                    print(f"   - {campo}: {valor[0][1][0]}")
                print()
            else:
                print(f"❌ Erro ao atualizar {login_ad}: {conn.result}\n")
        else:
            print(f"⚠️ Usuário {login_ad} não encontrado no AD.\n")

    conn.unbind()
else:
    print("❌ Falha ao conectar no Active Directory.")