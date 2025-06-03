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

        # Buscar o usuário pelo sAMAccountName
        conn.search(
            BASE_DN,
            f"(&(objectClass=user)(sAMAccountName={login_ad})(!(objectCategory=computer)))",
            attributes=["cn", "sAMAccountName", "userAccountControl"]
        )

        if conn.entries:
            usuario = conn.entries[0]
            print(f"🔒 Desativando {login_ad} - {usuario.cn.value}")

            # Define o valor para desativar a conta
            desativar = {
                'userAccountControl': [(ldap3.MODIFY_REPLACE, [514])]
            }

            if conn.modify(usuario.entry_dn, desativar):
                print(f"✅ Usuário {login_ad} desativado com sucesso!\n")
            else:
                print(f"❌ Erro ao desativar {login_ad}: {conn.result}\n")
        else:
            print(f"⚠️ Usuário {login_ad} não encontrado no AD.\n")

    conn.unbind()
else:
    print("❌ Falha ao conectar no Active Directory.")
