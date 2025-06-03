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
caminho_arquivo = os.path.join(os.path.expanduser("~"), "Downloads", "EmailxUsuarios.xlsx")

# Carregar a planilha
df = pd.read_excel(caminho_arquivo)

# Criar conexão com o AD
server = ldap3.Server(SERVIDOR_AD, get_info=ldap3.ALL)
conn = ldap3.Connection(server, user=USUARIO_AD, password=SENHA_AD, auto_bind=True)

if conn.bind():
    print("✅ Conexão com o Active Directory estabelecida com sucesso!\n")

    for index, row in df.iterrows():
        login_ad = str(row['Login AD']).strip()
        dn_origem = str(row["DN Origem"]).strip()
        dn_destino = str(row["DN Destino"]).strip()

        conn.search(
            BASE_DN,
            f"(&(objectClass=user)(sAMAccountName={login_ad})(!(objectCategory=computer)))",
            attributes=["cn", "sAMAccountName"]
        )

        if conn.entries:
            usuario = conn.entries[0]
            print(f"📁 Verificando movimentação do usuário {login_ad} - {usuario.cn.value}")

            if dn_origem.lower() != dn_destino.lower():
                if conn.modify_dn(dn_origem, f"CN={usuario.cn.value}", new_superior=dn_destino):
                    print(f"✅ Usuário {login_ad} movido com sucesso para {dn_destino}!\n")
                else:
                    print(f"❌ Erro ao mover {login_ad} para nova OU: {conn.result}\n")
            else:
                print(f"📍 Usuário {login_ad} já está na OU correta.\n")
        else:
            print(f"⚠️ Usuário {login_ad} não encontrado no AD.\n")

    conn.unbind()
else:
    print("❌ Falha ao conectar no Active Directory.")
