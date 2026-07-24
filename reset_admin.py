import os
import sqlite3
import hashlib

# Senha que será definida: admin
senha_hash = hashlib.sha256(b"admin").hexdigest()

# Locais onde o banco pode estar
candidatos = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "catalogo_feeu.db"),
    r"J:\Catalogo de Livros FEEU 7.27\catalogo_feeu.db",
]

encontrou = False

for db in candidatos:
    if not os.path.exists(db):
        continue

    encontrou = True

    print("=" * 60)
    print("Banco encontrado:")
    print(db)
    print("=" * 60)

    conn = sqlite3.connect(db)
    cur = conn.cursor()

    # Verifica se admin existe
    cur.execute("SELECT COUNT(*) FROM usuarios WHERE nome='admin'")
    existe = cur.fetchone()[0]

    if existe:
        cur.execute(
            "UPDATE usuarios SET senha=? WHERE nome='admin'",
            (senha_hash,)
        )
        print("Usuário admin já existia.")
    else:
        cur.execute(
            "INSERT INTO usuarios (nome, senha) VALUES ('admin', ?)",
            (senha_hash,)
        )
        print("Usuário admin não existia e foi criado.")

    # Ativa todas as permissões existentes para admin
    cur.execute("PRAGMA table_info(usuarios)")
    colunas = [r[1] for r in cur.fetchall()]

    for col in colunas:
        if col.startswith("perm_"):
            try:
                cur.execute(
                    f"UPDATE usuarios SET {col}=1 WHERE nome='admin'"
                )
            except Exception as e:
                print(f"Aviso: não foi possível alterar {col}: {e}")

    conn.commit()
    conn.close()

    print("Senha do admin redefinida para: admin")
    print("Permissões do admin ativadas.")
    print()

if not encontrou:
    print("Nenhum banco catalogo_feeu.db foi encontrado.")
    print("Verifique se o arquivo está na mesma pasta deste script.")

print("=" * 60)
print("Finalizado.")
print("Agora tente entrar com:")
print("Usuário: admin")
print("Senha: admin")
print("=" * 60)

input("Pressione Enter para fechar...")