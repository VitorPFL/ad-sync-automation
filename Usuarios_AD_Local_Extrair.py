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

# Atributos que queremos buscar
ATRIBUTOS = [
    "cn", "sAMAccountName", "mail", "userPrincipalName", "givenName",
    "sn", "telephoneNumber", "department", "title", "memberOf", "accountExpires", "Manager", "DisplayName",
    "physicalDeliveryOfficeName","Company","objectGUID", "distinguishedName","whenCreated", "employeeNumber",
    "description","userAccountControl","employeeID","employeeType"

]


# Função para converter DistinguishedName em formato Canonical
def dn_para_canonical(dn, dominio):
    try:
        partes_dn = dn.replace("CN=", "").replace("OU=", "").replace("DC=", "").split(",")
        caminho_objeto = "/".join(partes_dn[:-2])  # remove os dois últimos DCs (domínio)
        return f"{dominio}/{caminho_objeto}"
    except Exception:
        return ""

# Criar conexão com o AD
server = ldap3.Server(SERVIDOR_AD, get_info=ldap3.ALL)
conn = ldap3.Connection(server, user=USUARIO_AD, password=SENHA_AD, auto_bind=True)

# Verificar se a conexão foi bem-sucedida
if conn.bind():
    print("Conexão com o Active Directory estabelecida com sucesso!")

    # Consultar todos os usuários
    conn.search(
        BASE_DN,
        "(&(objectClass=user)(!(objectCategory=computer))(!(userAccountControl:1.2.840.113556.1.4.803:=514)))",
        attributes=ATRIBUTOS
    )

    # Processar os resultados
    usuarios = []
    for entry in conn.entries:
        dn = entry.distinguishedName.value if entry.distinguishedName else ""
        canonical = dn_para_canonical(dn, "grupomonto.local")
        usuarios.append({
            "Nome Completo": entry.cn.value,
            "Primeiro Nome": entry.givenName.value,
            "Sobrenome": entry.sn.value,
            "Nome Exibição": entry.displayName.value,
            "Login": entry.sAMAccountName.value,
            "Guid": entry.objectGUID.value,
            "E-mail": entry.mail.value if entry.mail else "",
            "Login Principal": entry.userPrincipalName.value if entry.userPrincipalName else "",
            "Telefone": entry.telephoneNumber.value if entry.telephoneNumber else "",
            "Departamento": entry.department.value if entry.department else "",
            "Cargo": entry.title.value if entry.title else "",
            "Supervisor": entry.Manager.value if entry.title else "",
            "Empresa": entry.Company.value if entry.title else "",
            "Local Fisico": entry.physicalDeliveryOfficeName.value if entry.title else "",
            "DistinguishedName": dn,
            "CanonicalName": canonical,
            "Data Criacao": entry.whenCreated.value if entry.whenCreated else "",
            "RE": entry.employeeID.value if entry.employeeID else "",
            "CPF": str(entry.employeeNumber.value).zfill(11) if entry.employeeNumber else "",
            "Status Ativo": "Sim" if entry.userAccountControl.value & 2 == 0 else "Não",
            "Descrição": entry.employeeType.value if entry.employeeType else ""
            #"Membro de Grupos": ", ".join(entry.memberOf) if entry.memberOf else "",
            # "Conta Expira": entry.accountExpires.value if entry.accountExpires else "Nunca"
        })

    # Criar um DataFrame e salvar em Excel
    df = pd.DataFrame(usuarios)

    # Remover fuso horário de colunas datetime, se existirem
    for col in df.select_dtypes(include=['datetime64[ns, UTC]']).columns:
        df[col] = df[col].dt.tz_localize(None)

    # Definir o caminho do arquivo
    caminho_saida = r"C:\Users\vitor.leite\GRUPO MONTO\TI - BI\05 TI\01 Bases\Usuarios_AD.xlsx"

    # Salvar no Excel
    df.to_excel(caminho_saida, index=False)

    print(f"Arquivo salvo com sucesso: {caminho_saida}")

    # else:
    # print("Falha ao conectar no Active Directory!")

# Fechar conexão
conn.unbind()
