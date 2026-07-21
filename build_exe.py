import subprocess
import sys
import os

def build():
    print("Iniciando empacotamento do Gerador de Certificados com PyInstaller...")
    
    # Configura o comando do PyInstaller
    # --clean: Limpa o cache do PyInstaller antes do build
    # --onefile: Empacota tudo em um único arquivo .exe
    # --noconsole: Não exibe a janela preta do prompt de comando quando o executável roda
    # --name: Nome final do arquivo executável
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--clean",
        "--onefile",
        "--noconsole",
        "--name=GeradorCertificados",
        "main.py"
    ]
    
    try:
        # Executa o PyInstaller
        subprocess.run(cmd, check=True)
        print("\nEmpacotamento concluído com sucesso!")
        print(f"O executável autônomo .exe foi gerado na pasta: {os.path.abspath('dist')}")
    except subprocess.CalledProcessError as e:
        print(f"\nErro durante o empacotamento: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    build()
