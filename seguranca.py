import hashlib
import datetime
import getpass
import os

def gerar_hash_arquivo(caminho_arquivo):
    """Gera o hash SHA-256 lendo o arquivo em blocos."""
    sha256_hash = hashlib.sha256()
    try:
        with open(caminho_arquivo, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except Exception as e:
        print(f"Erro ao gerar hash: {e}")
        return None

def coletar_dados_auditoria():
    """Captura o usuário do Windows e o horário exato."""
    usuario = getpass.getuser()
    nome_maquina = os.environ.get('COMPUTERNAME', 'Maquina')
    data_hora_atual = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    return {
        "usuario_sistema": f"{nome_maquina}\\{usuario}",
        "data_hora": data_hora_atual
    }