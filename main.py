import sys
from PySide6.QtWidgets import QApplication
from gui import CertificateGeneratorApp

def main():
    # Cria a instância da aplicação do Qt
    app = QApplication(sys.argv)
    
    # Cria e exibe a janela principal
    window = CertificateGeneratorApp()
    window.show()
    
    # Executa o loop principal da interface
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
