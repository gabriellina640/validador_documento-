import sqlite3
import os

NOME_BANCO = "auditoria_corregedoria.db"

def inicializar_banco():
    conexao = sqlite3.connect(NOME_BANCO)
    cursor = conexao.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS registro_auditoria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_original TEXT NOT NULL,
            caminho_processado TEXT NOT NULL,
            hash_sha256 TEXT NOT NULL,
            usuario_sistema TEXT NOT NULL,
            data_hora TEXT NOT NULL
        )
    ''')
    conexao.commit()
    conexao.close()

def registrar_documento(nome_original, caminho_processado, hash_sha256, usuario, data_hora):
    try:
        conexao = sqlite3.connect(NOME_BANCO)
        cursor = conexao.cursor()
        cursor.execute('''
            INSERT INTO registro_auditoria 
            (nome_original, caminho_processado, hash_sha256, usuario_sistema, data_hora)
            VALUES (?, ?, ?, ?, ?)
        ''', (nome_original, caminho_processado, hash_sha256, usuario, data_hora))
        conexao.commit()
        conexao.close()
        return True
    except Exception as e:
        print(f"Erro BD: {e}")
        return False

def buscar_por_hash(hash_busca):
    try:
        if not os.path.exists(NOME_BANCO): return None
        conexao = sqlite3.connect(NOME_BANCO)
        cursor = conexao.cursor()
        cursor.execute('SELECT nome_original, data_hora, usuario_sistema FROM registro_auditoria WHERE hash_sha256 = ?', (hash_busca,))
        resultado = cursor.fetchone()
        conexao.close()
        return resultado
    except:
        return None