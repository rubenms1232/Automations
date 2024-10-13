import os
import subprocess
import sys

# Verifica se o pip está instalado
def check_pip():
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "--version"])
        print("Pip já está instalado.")
    except subprocess.CalledProcessError:
        print("Pip não está instalado. Instalando o pip...")
        subprocess.check_call([sys.executable, "-m", "ensurepip", "--default-pip"])

# Instala os pacotes necessários
def install_requirements():
    print("Instalando pacotes necessários...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])

if __name__ == "__main__":
    check_pip()
    install_requirements()
    print("Instalação completa!")
