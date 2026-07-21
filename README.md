# Gerador de Certificados 

Este é um projeto desenvolvido em Python utilizando a biblioteca **PySide6 (Qt)** e automação COM do PowerPoint (`pywin32`) para gerar certificados personalizados em lote a partir de uma planilha do Excel (`.xlsx`, `.xls` ou `.csv`) e um arquivo de apresentação do PowerPoint (`.pptx`) como modelo.

## Funcionalidades

- **Mapeamento Flexível**: Identifica automaticamente marcadores no slide (ex: `#nome`, `#turma`) e permite associá-los de forma dinâmica a qualquer coluna da planilha.
- **Auto-Mapeamento**: Botão para dar match inteligente entre placeholders e cabeçalhos de coluna equivalentes.
- **Formatos de Saída**: Permite escolher salvar os certificados gerados como **PDF** ou **PNG**.
- **Nome de Arquivo Customizado**: Permite configurar um padrão de nome dinâmico para os arquivos salvos usando chaves, por exemplo: `Certificado_{Nome}_{Turma}`.
- **Operação Assíncrona (Multithreaded)**: O processamento em lote ocorre em segundo plano, evitando que a interface trave e permitindo o cancelamento a qualquer momento.
- **Gerador de Instalador**: Inclui script para compilar o programa em um executável (.exe) de arquivo único.

---

##  Pré-requisitos

1. **Sistema Operacional**: Windows.
2. **Microsoft PowerPoint**: Como o gerador utiliza o motor oficial do PowerPoint para manter fidelidade visual total (fontes, cores, alinhamentos), **é obrigatório ter o Microsoft Office/PowerPoint instalado** na máquina que executará o script ou executável.
3. **Python 3.8+** (Recomendado Python 3.11.x) - *Nota: Necessário apenas se for rodar o código fonte. Se for rodar pelo arquivo executável `.exe`, o Python **NÃO** é necessário.*

---

##  Instalação e Execução

### Modo Rápido (Via Executável)
Se você já tem o arquivo executável compilado, **basta dar dois cliques em `GeradorCertificados.exe` na pasta `dist/`** para abrir o programa imediatamente.

### Modo Desenvolvedor (Via Código Fonte)
Caso queira rodar o código fonte diretamente:

1. **Instalar as Dependências**:
   Abra o terminal/Prompt de Comando na pasta do projeto e execute:
   ```bash
   pip install -r requirements.txt
   ```

2. **Executar o Programa**:
   ```bash
   python main.py
   ```

---

## Como Gerar o Executável (.exe)

Se você deseja gerar um instalador/arquivo único executável para distribuir para outras pessoas que não possuem Python instalado, use o script `build_exe.py`:

```bash
python build_exe.py
```

O PyInstaller empacotará o projeto e criará um executável na pasta **`dist/GeradorCertificados.exe`**.

*Nota: O usuário final do executável `.exe` ainda precisará do Microsoft PowerPoint instalado na máquina para que a geração de arquivos funcione.*

---

## Como Usar (Passo a Passo)

1. **Prepare o Modelo no PowerPoint (`.pptx`)**:
   - Crie o slide do seu certificado.
   - Onde você quer que as informações da planilha apareçam, insira caixas de texto com o padrão `#nomedocampo` (ex: `#nome`, `#turma`, `#data`).
   
2. **Prepare a Planilha (`.xlsx` ou `.csv`)**:
   - Crie colunas contendo as informações (ex: coluna `Nome`, coluna `Turma`).
   - Cada linha representará uma pessoa que receberá o certificado.

3. **Configure no Programa**:
   - Abra o Gerador.
   - Selecione o arquivo `.pptx`.
   - Selecione a planilha Excel/CSV.
   - Escolha o número do slide que corresponde ao modelo (caso o PowerPoint tenha múltiplos slides).
   - No painel de **Mapeamento de Informações**, associe cada `#marcador` à coluna correspondente (use o botão *Auto Mapear Colunas* para agilizar).
   - Defina a pasta onde os certificados serão salvos.
   - Selecione o formato de saída (PDF ou PNG).
   - Personalize o padrão de nome (ex: `Certificado_{Nome}`).
   
4. **Gere**:
   - Clique em **GERAR CERTIFICADOS**.
   - Acompanhe a barra de progresso. Ao final, o programa perguntará se deseja abrir a pasta contendo todos os certificados gerados.

