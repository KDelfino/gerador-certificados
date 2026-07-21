import os
import re
import win32com.client

def limpar_nome_arquivo(nome):
    """
    Remove caracteres inválidos para nomes de arquivos no Windows.
    """
    return re.sub(r'[\\/*?:"<>|]', "", nome).strip()

def substituir_texto_shapes(shapes, texto_antigo, texto_novo):
    """
    Percorre as formas do slide para substituir o texto.
    A recursão garante que textos dentro de elementos agrupados também sejam lidos e alterados.
    """
    for shape in shapes:
        # Verifica se é um grupo de formas (msoGroup = 6)
        if shape.Type == 6:
            substituir_texto_shapes(shape.GroupItems, texto_antigo, texto_novo)
        elif shape.HasTextFrame:
            if shape.TextFrame.HasText:
                text_range = shape.TextFrame.TextRange
                # O método Replace preserva a formatação da fonte (tamanho, cor, estilo)
                if texto_antigo in text_range.Text:
                    text_range.Replace(texto_antigo, texto_novo)

def gerar_certificados():
    # Configuração de caminhos baseada no diretório onde este script está salvo
    diretorio_base = os.path.dirname(os.path.abspath(__file__))
    pasta_arquivos = os.path.join(diretorio_base, "Arquivos")
    pasta_certificados = os.path.join(diretorio_base, "certificados")

    arquivo_pessoas = os.path.join(pasta_arquivos, "pessoas.txt")
    arquivo_modelo = os.path.join(pasta_arquivos, "modelo.pptx")

    # Verificações de existência dos arquivos de entrada
    if not os.path.exists(arquivo_pessoas):
        print(f"Erro: Arquivo '{arquivo_pessoas}' não encontrado.")
        return
    if not os.path.exists(arquivo_modelo):
        print(f"Erro: Arquivo '{arquivo_modelo}' não encontrado.")
        return

    # 1. Criar automaticamente a pasta "certificados" caso ela não exista
    if not os.path.exists(pasta_certificados):
        os.makedirs(pasta_certificados)
        print(f"Pasta '{pasta_certificados}' criada com sucesso.\n")

    # 2. Ler todos os nomes do arquivo
    with open(arquivo_pessoas, 'r', encoding='utf-8') as f:
        nomes = [linha.strip() for linha in f if linha.strip()]

    if not nomes:
        print("Nenhum nome encontrado no arquivo pessoas.txt.")
        return

    print(f"Iniciando a geração de {len(nomes)} certificados...\n")

    # 3. Iniciar automação COM via pywin32
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    
    try:
        # Abre a apresentação
        presentation = powerpoint.Presentations.Open(arquivo_modelo, ReadOnly=True, WithWindow=False)

        # 4. Utilizar apenas o slide indicado (os índices no PowerPoint COM começam em 1)
        slide_index = 9
        slide = presentation.Slides(slide_index)

        # 5. Para cada nome, realizar a substituição e exportação
        for i, nome in enumerate(nomes, 1):
            nome_arquivo = f"{limpar_nome_arquivo(nome)}.png"
            caminho_saida = os.path.join(pasta_certificados, nome_arquivo)

            # Substitui o texto "#nome" pelo nome da pessoa
            substituir_texto_shapes(slide.Shapes, "#nome", nome)

            # Exporta o certificado como imagem PNG
            slide.Export(caminho_saida, "PNG")

            # Desfaz a substituição para retornar ao estado "#nome" e preparar a próxima iteração
            substituir_texto_shapes(slide.Shapes, nome, "#nome")

            # Exibe mensagem de progresso no terminal
            print(f"[{i}/{len(nomes)}] Certificado exportado: {nome_arquivo}")

        print("\nProcesso concluído com sucesso!")

    except Exception as e:
        print(f"\nOcorreu um erro durante a automação do PowerPoint: {e}")

    finally:
        # 6. Fechar corretamente os processos do PowerPoint ao final da execução
        try:
            presentation.Close()
        except:
            pass
        powerpoint.Quit()

if __name__ == "__main__":
    gerar_certificados()