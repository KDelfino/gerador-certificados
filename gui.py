import os
import re
from PySide6.QtCore import Qt, QThread, QObject, Signal, Slot
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QComboBox, QGroupBox, QGridLayout,
    QProgressBar, QPlainTextEdit, QScrollArea, QMessageBox, QRadioButton, QButtonGroup
)
from PySide6.QtGui import QIcon, QFont, QLinearGradient, QColor, QPalette

from engine import (
    ler_colunas_planilha, ler_dados_planilha,
    obter_quantidade_slides, escanear_placeholders_slide
)

# Folha de Estilos QSS Premium
ESTILO_CANDIDATO = """
QMainWindow {
    background-color: #1a1a1a;
}
QWidget {
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
    font-size: 13px;
    color: #e0e0e0;
}
QLabel {
    font-weight: 500;
}
QLineEdit {
    background-color: #2b2b2b;
    border: 1px solid #444444;
    border-radius: 6px;
    padding: 6px 10px;
    color: #ffffff;
}
QLineEdit:focus {
    border: 1px solid #007acc;
}
QPushButton {
    background-color: #007acc;
    color: #ffffff;
    border: none;
    border-radius: 6px;
    padding: 8px 16px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #0098ff;
}
QPushButton:pressed {
    background-color: #005999;
}
QPushButton:disabled {
    background-color: #333333;
    color: #777777;
}
QPushButton#btn_cancelar {
    background-color: #d32f2f;
}
QPushButton#btn_cancelar:hover {
    background-color: #f44336;
}
QPushButton#btn_cancelar:pressed {
    background-color: #b71c1c;
}
QPushButton.secundario {
    background-color: #2b2b2b;
    border: 1px solid #444444;
    color: #e0e0e0;
}
QPushButton.secundario:hover {
    background-color: #3d3d3d;
    border-color: #555555;
}
QPushButton.secundario:pressed {
    background-color: #222222;
}
QComboBox {
    background-color: #2b2b2b;
    border: 1px solid #444444;
    border-radius: 6px;
    padding: 5px 10px;
    color: #ffffff;
}
QComboBox:focus {
    border: 1px solid #007acc;
}
QComboBox::drop-down {
    border: none;
}
QComboBox QAbstractItemView {
    background-color: #2b2b2b;
    border: 1px solid #444444;
    selection-background-color: #007acc;
    selection-color: #ffffff;
}
QGroupBox {
    border: 1px solid #333333;
    border-radius: 8px;
    margin-top: 15px;
    padding-top: 15px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 0 5px;
    color: #007acc;
    font-weight: bold;
    font-size: 14px;
}
QProgressBar {
    border: 1px solid #333333;
    border-radius: 6px;
    text-align: center;
    background-color: #111111;
    font-weight: bold;
}
QProgressBar::chunk {
    background-color: #007acc;
    border-radius: 5px;
}
QPlainTextEdit {
    background-color: #111111;
    border: 1px solid #333333;
    border-radius: 6px;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
    color: #a5d6a7;
}
QScrollBar:vertical {
    border: none;
    background: #1a1a1a;
    width: 10px;
    margin: 0px;
}
QScrollBar::handle:vertical {
    background: #444444;
    min-height: 20px;
    border-radius: 5px;
}
QScrollBar::handle:vertical:hover {
    background: #555555;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    border: none;
    background: none;
}
"""

class GeneratorWorker(QObject):
    """
    Worker que executa a geração de certificados em uma thread secundária.
    """
    progress = Signal(int, int, str)
    finished = Signal()
    error = Signal(str)

    def __init__(self, pptx, slide_idx, data, mapping, fmt, out_dir, pattern):
        super().__init__()
        self.pptx = pptx
        self.slide_idx = slide_idx
        self.data = data
        self.mapping = mapping
        self.fmt = fmt
        self.out_dir = out_dir
        self.pattern = pattern
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def is_cancelled(self):
        return self._is_cancelled

    def run(self):
        try:
            from engine import gerar_certificados_job
            
            def progress_cb(current, total, msg):
                self.progress.emit(current, total, msg)
                
            gerar_certificados_job(
                caminho_pptx=self.pptx,
                slide_index=self.slide_idx,
                dados_planilha=self.data,
                mapeamento=self.mapping,
                formato_saida=self.fmt,
                pasta_destino=self.out_dir,
                padrao_nome=self.pattern,
                progress_callback=progress_cb,
                is_cancelled_callback=self.is_cancelled
            )
            
            if not self._is_cancelled:
                self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))


class CertificateGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Certificados Dinâmico")
        self.resize(750, 780)
        self.setStyleSheet(ESTILO_CANDIDATO)
        
        # Variáveis de Estado
        self.colunas_planilha = []
        self.dados_planilha = []
        self.placeholders_slide = []
        self.mapeamentos_widgets = {} # {placeholder: QComboBox}
        
        # Thread de geração
        self.thread_geracao = None
        self.worker_geracao = None
        
        self.init_ui()

    def init_ui(self):
        widget_central = QWidget()
        self.setCentralWidget(widget_central)
        layout_principal = QVBoxLayout(widget_central)
        layout_principal.setContentsMargins(20, 20, 20, 20)
        layout_principal.setSpacing(15)

        # 1. Cabeçalho Decorado
        widget_cabecalho = QWidget()
        widget_cabecalho.setFixedHeight(70)
        widget_cabecalho.setStyleSheet("""
            background-color: #1a237e;
            border-radius: 8px;
        """)
        layout_cabecalho = QVBoxLayout(widget_cabecalho)
        layout_cabecalho.setContentsMargins(15, 0, 15, 0)
        layout_cabecalho.setSpacing(2)
        
        lbl_titulo = QLabel("GERADOR DE CERTIFICADOS DINÂMICO")
        lbl_titulo.setFont(QFont("Segoe UI", 16, QFont.Bold))
        lbl_titulo.setStyleSheet("color: #ffffff; background: transparent;")
        
        lbl_subtitulo = QLabel("Gere certificados em PDF ou PNG mapeando planilhas e modelos PPTX.")
        lbl_subtitulo.setFont(QFont("Segoe UI", 10))
        lbl_subtitulo.setStyleSheet("color: #b0bec5; background: transparent;")
        
        layout_cabecalho.addWidget(lbl_titulo)
        layout_cabecalho.addWidget(lbl_subtitulo)
        layout_principal.addWidget(widget_cabecalho)

        # 2. Grupo: Arquivos de Entrada
        grupo_arquivos = QGroupBox("Arquivos de Entrada")
        layout_grupo_arq = QGridLayout(grupo_arquivos)
        layout_grupo_arq.setSpacing(10)
        
        # Modelo PPTX
        layout_grupo_arq.addWidget(QLabel("Modelo PowerPoint (.pptx):"), 0, 0)
        self.txt_modelo_pptx = QLineEdit()
        self.txt_modelo_pptx.setPlaceholderText("Selecione o arquivo .pptx...")
        self.txt_modelo_pptx.textChanged.connect(self.ao_alterar_modelo_pptx)
        layout_grupo_arq.addWidget(self.txt_modelo_pptx, 0, 1)
        btn_pptx = QPushButton("Buscar")
        btn_pptx.setProperty("class", "secundario")
        btn_pptx.clicked.connect(self.buscar_modelo_pptx)
        layout_grupo_arq.addWidget(btn_pptx, 0, 2)
        
        # Planilha Excel
        layout_grupo_arq.addWidget(QLabel("Planilha Excel ou CSV:"), 1, 0)
        self.txt_planilha = QLineEdit()
        self.txt_planilha.setPlaceholderText("Selecione o arquivo .xlsx, .xls ou .csv...")
        self.txt_planilha.textChanged.connect(self.ao_alterar_planilha)
        layout_grupo_arq.addWidget(self.txt_planilha, 1, 1)
        btn_planilha = QPushButton("Buscar")
        btn_planilha.setProperty("class", "secundario")
        btn_planilha.clicked.connect(self.buscar_planilha)
        layout_grupo_arq.addWidget(btn_planilha, 1, 2)
        
        # Seleção de Slide do Modelo
        layout_grupo_arq.addWidget(QLabel("Slide do Certificado:"), 2, 0)
        self.cb_slides = QComboBox()
        self.cb_slides.setEnabled(False)
        self.cb_slides.currentIndexChanged.connect(self.ao_alterar_slide)
        layout_grupo_arq.addWidget(self.cb_slides, 2, 1, 1, 2)
        
        layout_principal.addWidget(grupo_arquivos)

        # 3. Grupo: Mapeamento de Placeholders
        self.grupo_mapeamento = QGroupBox("Mapeamento de Informações")
        layout_grupo_map = QVBoxLayout(self.grupo_mapeamento)
        
        # Container com Scroll para o Mapeamento Dinâmico
        self.scroll_mapeamento = QScrollArea()
        self.scroll_mapeamento.setWidgetResizable(True)
        self.scroll_mapeamento.setMinimumHeight(220)
        self.scroll_mapeamento.setStyleSheet("background-color: #222222; border: none; border-radius: 6px;")
        
        self.widget_conteudo_map = QWidget()
        self.layout_conteudo_map = QGridLayout(self.widget_conteudo_map)
        self.layout_conteudo_map.setSpacing(10)
        self.scroll_mapeamento.setWidget(self.widget_conteudo_map)
        
        layout_grupo_map.addWidget(self.scroll_mapeamento)
        
        # Botão de auto-mapeamento rápido
        self.btn_auto_map = QPushButton("Auto Mapear Colunas")
        self.btn_auto_map.setProperty("class", "secundario")
        self.btn_auto_map.clicked.connect(self.auto_mapear_colunas)
        self.btn_auto_map.setEnabled(False)
        layout_grupo_map.addWidget(self.btn_auto_map)
        
        layout_principal.addWidget(self.grupo_mapeamento)

        # 4. Grupo: Configurações de Saída
        grupo_saida = QGroupBox("Configurações de Saída")
        layout_grupo_saida = QGridLayout(grupo_saida)
        layout_grupo_saida.setSpacing(10)
        
        # Pasta de Destino
        layout_grupo_saida.addWidget(QLabel("Pasta de Destino:"), 0, 0)
        self.txt_pasta_destino = QLineEdit()
        self.txt_pasta_destino.setPlaceholderText("Selecione onde salvar os certificados...")
        layout_grupo_saida.addWidget(self.txt_pasta_destino, 0, 1)
        btn_destino = QPushButton("Buscar")
        btn_destino.setProperty("class", "secundario")
        btn_destino.clicked.connect(self.buscar_pasta_destino)
        layout_grupo_saida.addWidget(btn_destino, 0, 2)
        
        # Formato de Saída (PDF ou PNG)
        layout_grupo_saida.addWidget(QLabel("Formato de Saída:"), 1, 0)
        layout_formatos = QHBoxLayout()
        self.rb_pdf = QRadioButton("PDF (.pdf)")
        self.rb_pdf.setChecked(True)
        self.rb_png = QRadioButton("PNG (.png)")
        
        self.grupo_formatos = QButtonGroup()
        self.grupo_formatos.addButton(self.rb_pdf)
        self.grupo_formatos.addButton(self.rb_png)
        
        layout_formatos.addWidget(self.rb_pdf)
        layout_formatos.addWidget(self.rb_png)
        layout_formatos.addStretch()
        layout_grupo_saida.addLayout(layout_formatos, 1, 1, 1, 2)
        
        # Padrão do Nome do Arquivo
        layout_grupo_saida.addWidget(QLabel("Padrão do Nome:"), 2, 0)
        self.txt_padrao_nome = QLineEdit()
        self.txt_padrao_nome.setText("Certificado_{Nome}")
        self.txt_padrao_nome.setToolTip("Substitua colunas da planilha usando chaves, ex: Certificado_{Nome}_{Turma}")
        layout_grupo_saida.addWidget(self.txt_padrao_nome, 2, 1, 1, 2)
        
        self.lbl_dica_nome = QLabel("Disponível: (Carregue uma planilha primeiro)")
        self.lbl_dica_nome.setFont(QFont("Segoe UI", 9, QFont.Normal, True))
        self.lbl_dica_nome.setStyleSheet("color: #888888;")
        layout_grupo_saida.addWidget(self.lbl_dica_nome, 3, 1, 1, 2)
        
        layout_principal.addWidget(grupo_saida)

        # 5. Seção de Geração e Log
        layout_botoes_gerar = QHBoxLayout()
        self.btn_gerar = QPushButton("GERAR CERTIFICADOS")
        self.btn_gerar.setFixedHeight(40)
        self.btn_gerar.setStyleSheet("""
            background-color: #2ecc71;
            font-size: 14px;
        """)
        self.btn_gerar.clicked.connect(self.iniciar_geracao)
        
        self.btn_cancelar = QPushButton("CANCELAR")
        self.btn_cancelar.setObjectName("btn_cancelar")
        self.btn_cancelar.setFixedHeight(40)
        self.btn_cancelar.setEnabled(False)
        self.btn_cancelar.clicked.connect(self.cancelar_geracao)
        
        layout_botoes_gerar.addWidget(self.btn_gerar, 3)
        layout_botoes_gerar.addWidget(self.btn_cancelar, 1)
        layout_principal.addLayout(layout_botoes_gerar)

        # Barra de Progresso
        self.bar_progresso = QProgressBar()
        self.bar_progresso.setValue(0)
        self.bar_progresso.setFixedHeight(22)
        self.bar_progresso.setVisible(False)
        layout_principal.addWidget(self.bar_progresso)

        # Log Console
        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setPlaceholderText("Os logs do processo de geração aparecerão aqui...")
        self.txt_log.setFixedHeight(120)
        layout_principal.addWidget(self.txt_log)

    # --- AÇÕES DO USUÁRIO ---
    
    def buscar_modelo_pptx(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Modelo PowerPoint", "", "Apresentações PowerPoint (*.pptx)"
        )
        if caminho:
            self.txt_modelo_pptx.setText(os.path.abspath(caminho))

    def buscar_planilha(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Planilha ou CSV", "", "Planilhas (*.xlsx *.xls *.csv)"
        )
        if caminho:
            self.txt_planilha.setText(os.path.abspath(caminho))

    def buscar_pasta_destino(self):
        caminho = QFileDialog.getExistingDirectory(
            self, "Selecionar Pasta de Destino"
        )
        if caminho:
            self.txt_pasta_destino.setText(os.path.abspath(caminho))

    # --- LÓGICA DE DADOS & COMPONENTES ---

    def ao_alterar_modelo_pptx(self, caminho):
        if not caminho or not os.path.exists(caminho):
            self.cb_slides.clear()
            self.cb_slides.setEnabled(False)
            self.placeholders_slide = []
            self.limpar_mapeamentos()
            return

        # Busca quantidade de slides
        self.txt_log.appendPlainText("Carregando apresentação PowerPoint...")
        self.setCursor(Qt.WaitCursor)
        try:
            total_slides = obter_quantidade_slides(caminho)
            self.cb_slides.clear()
            self.cb_slides.setEnabled(True)
            for i in range(1, total_slides + 1):
                self.cb_slides.addItem(f"Slide {i}", i)
            
            # Se for carregado com sucesso, tenta escanear o primeiro
            self.ao_alterar_slide(0)
            self.txt_log.appendPlainText(f"PowerPoint carregado. Total de slides: {total_slides}")
        except Exception as e:
            self.txt_log.appendPlainText(f"Erro ao ler modelo PowerPoint: {e}")
            QMessageBox.critical(self, "Erro", f"Não foi possível ler o arquivo PowerPoint.\nCertifique-se de que ele não está aberto ou corrompido.\n\nDetalhes: {e}")
            self.txt_modelo_pptx.clear()
        finally:
            self.unsetCursor()

    def ao_alterar_planilha(self, caminho):
        if not caminho or not os.path.exists(caminho):
            self.colunas_planilha = []
            self.dados_planilha = []
            self.lbl_dica_nome.setText("Disponível: (Carregue uma planilha primeiro)")
            self.atualizar_comboboxes_mapeamento()
            self.btn_auto_map.setEnabled(False)
            return

        self.txt_log.appendPlainText("Lendo planilha...")
        try:
            self.colunas_planilha = ler_colunas_planilha(caminho)
            self.dados_planilha = ler_dados_planilha(caminho)
            
            # Atualiza dicas
            colunas_str = ", ".join([f"{{{col}}}" for col in self.colunas_planilha])
            self.lbl_dica_nome.setText(f"Disponível: {colunas_str}")
            
            self.atualizar_comboboxes_mapeamento()
            self.btn_auto_map.setEnabled(len(self.colunas_planilha) > 0 and len(self.placeholders_slide) > 0)
            if len(self.colunas_planilha) > 0 and len(self.placeholders_slide) > 0:
                self.auto_mapear_colunas()
            
            self.txt_log.appendPlainText(f"Planilha carregada. Colunas encontradas: {len(self.colunas_planilha)}. Linhas de dados: {len(self.dados_planilha)}")
            
            # Tenta preencher automaticamente a pasta de destino caso esteja vazia
            if not self.txt_pasta_destino.text():
                diretorio_planilha = os.path.dirname(caminho)
                pasta_padrao = os.path.join(diretorio_planilha, "certificados")
                self.txt_pasta_destino.setText(pasta_padrao)
                
        except Exception as e:
            self.txt_log.appendPlainText(f"Erro ao ler planilha: {e}")
            QMessageBox.critical(self, "Erro", f"Não foi possível ler a planilha.\n\nDetalhes: {e}")
            self.txt_planilha.clear()

    def ao_alterar_slide(self, index):
        if index < 0 or not self.txt_modelo_pptx.text():
            self.placeholders_slide = []
            self.limpar_mapeamentos()
            return
            
        slide_idx = self.cb_slides.itemData(index)
        self.txt_log.appendPlainText(f"Escaneando placeholders no Slide {slide_idx}...")
        self.setCursor(Qt.WaitCursor)
        try:
            self.placeholders_slide = escanear_placeholders_slide(self.txt_modelo_pptx.text(), slide_idx)
            self.gerar_grade_mapeamento()
            self.btn_auto_map.setEnabled(len(self.colunas_planilha) > 0 and len(self.placeholders_slide) > 0)
            self.txt_log.appendPlainText(f"Placeholders encontrados no slide: {', '.join(self.placeholders_slide) if self.placeholders_slide else 'Nenhum'}")
        except Exception as e:
            self.txt_log.appendPlainText(f"Erro ao escanear slide: {e}")
            QMessageBox.warning(self, "Aviso", f"Erro ao ler placeholders do slide selecionado.\n\nDetalhes: {e}")
        finally:
            self.unsetCursor()

    def limpar_mapeamentos(self):
        # Remove todos os widgets antigos do layout de mapeamento
        while self.layout_conteudo_map.count():
            item = self.layout_conteudo_map.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.mapeamentos_widgets.clear()

    def gerar_grade_mapeamento(self):
        self.limpar_mapeamentos()
        
        if not self.placeholders_slide:
            lbl_aviso = QLabel("Nenhum marcador (#texto) encontrado neste slide.")
            lbl_aviso.setStyleSheet("color: #888888; font-style: italic;")
            self.layout_conteudo_map.addWidget(lbl_aviso, 0, 0, 1, 2)
            return

        # Cabeçalhos da Grade
        lbl_h1 = QLabel("Marcador no Slide")
        lbl_h1.setStyleSheet("font-weight: bold; color: #007acc; font-size: 15px; padding-bottom: 5px;")
        lbl_h2 = QLabel("Coluna Correspondente na Planilha")
        lbl_h2.setStyleSheet("font-weight: bold; color: #007acc; font-size: 15px; padding-bottom: 5px;")
        
        self.layout_conteudo_map.addWidget(lbl_h1, 0, 0)
        self.layout_conteudo_map.addWidget(lbl_h2, 0, 1)
        self.layout_conteudo_map.setVerticalSpacing(12)
        self.layout_conteudo_map.setHorizontalSpacing(20)

        # Adiciona uma linha para cada placeholder
        for idx, placeholder in enumerate(self.placeholders_slide, 1):
            lbl_placeholder = QLabel(placeholder)
            lbl_placeholder.setStyleSheet("font-family: monospace; font-size: 16px; font-weight: bold; color: #ffffff; padding: 4px;")
            
            cb_colunas = QComboBox()
            cb_colunas.setStyleSheet("font-size: 14px; padding: 4px 8px; min-height: 28px;")
            self.popular_combobox_colunas(cb_colunas)
            
            self.layout_conteudo_map.addWidget(lbl_placeholder, idx, 0)
            self.layout_conteudo_map.addWidget(cb_colunas, idx, 1)
            
            self.mapeamentos_widgets[placeholder] = cb_colunas
            
        # Tenta mapear automaticamente
        self.auto_mapear_colunas()

    def popular_combobox_colunas(self, combobox):
        combobox.clear()
        combobox.addItem("[Ignorar / Não substituir]", None)
        for col in self.colunas_planilha:
            combobox.addItem(col, col)

    def atualizar_comboboxes_mapeamento(self):
        for placeholder, cb in self.mapeamentos_widgets.items():
            valor_atual = cb.currentData()
            self.popular_combobox_colunas(cb)
            
            # Tenta re-selecionar o valor anterior se ele ainda existir na lista
            if valor_atual:
                idx = cb.findData(valor_atual)
                if idx >= 0:
                    cb.setCurrentIndex(idx)

    @Slot()
    def auto_mapear_colunas(self):
        """
        Mapeia de forma inteligente as colunas da planilha aos placeholders baseando-se no nome.
        Ignora diferença de caixa (case-insensitive) e caracteres como '#' e '_'.
        """
        print(f"[DIAGNÓSTICO] auto_mapear_colunas chamado!")
        print(f"[DIAGNÓSTICO] Colunas planilha: {self.colunas_planilha}")
        print(f"[DIAGNÓSTICO] Placeholders slide: {self.placeholders_slide}")
        
        if not self.colunas_planilha or not self.placeholders_slide:
            print("[DIAGNÓSTICO] Cancelado: colunas ou placeholders vazios.")
            return

        for placeholder, cb in self.mapeamentos_widgets.items():
            # Limpa o placeholder para tentar dar match (ex: "#nome_completo" -> "nomecompleto")
            p_limpo = re.sub(r'[^a-zA-Z0-9]', '', placeholder).lower()
            
            match_index = -1
            print(f"[DIAGNÓSTICO] Tentando mapear placeholder '{placeholder}' (limpo: '{p_limpo}')")
            
            # Busca pelo melhor match
            for idx in range(1, cb.count()): # Ignora o índice 0 que é o "[Ignorar]"
                col_text = cb.itemText(idx)
                col_limpa = re.sub(r'[^a-zA-Z0-9]', '', col_text).lower()
                print(f"[DIAGNÓSTICO]   Comparando com coluna '{col_text}' (limpa: '{col_limpa}')")
                
                # Se for match perfeito ou se uma contiver a outra
                if p_limpo == col_limpa or p_limpo in col_limpa or col_limpa in p_limpo:
                    match_index = idx
                    print(f"[DIAGNÓSTICO]     MATCH ENCONTRADO na coluna '{col_text}' (index {idx})")
                    break
                    
            if match_index >= 0:
                cb.setCurrentIndex(match_index)
                print(f"[DIAGNÓSTICO] Mapeado com sucesso: {placeholder} -> {cb.itemText(match_index)}")
            else:
                print(f"[DIAGNÓSTICO] Não mapeado: {placeholder}")

    # --- LÓGICA DE THREAD DE PROCESSAMENTO ---

    def alternar_estado_gui(self, rodando):
        self.txt_modelo_pptx.setEnabled(not rodando)
        self.txt_planilha.setEnabled(not rodando)
        self.cb_slides.setEnabled(not rodando and self.cb_slides.count() > 0)
        self.scroll_mapeamento.setEnabled(not rodando)
        self.btn_auto_map.setEnabled(not rodando and len(self.colunas_planilha) > 0 and len(self.placeholders_slide) > 0)
        self.txt_pasta_destino.setEnabled(not rodando)
        self.rb_pdf.setEnabled(not rodando)
        self.rb_png.setEnabled(not rodando)
        self.txt_padrao_nome.setEnabled(not rodando)
        
        # Botões de controle
        self.btn_gerar.setEnabled(not rodando)
        if rodando:
            self.btn_gerar.setText("GERANDO CERTIFICADOS...")
        else:
            self.btn_gerar.setText("GERAR CERTIFICADOS")
        self.btn_cancelar.setEnabled(rodando)
        self.bar_progresso.setVisible(rodando)

    @Slot()
    def iniciar_geracao(self):
        # 1. Validações básicas
        pptx_path = self.txt_modelo_pptx.text()
        sheet_path = self.txt_planilha.text()
        out_dir = self.txt_pasta_destino.text()
        pattern = self.txt_padrao_nome.text()
        
        if not pptx_path or not os.path.exists(pptx_path):
            QMessageBox.warning(self, "Aviso", "Por favor, selecione um arquivo de modelo PowerPoint (.pptx) válido.")
            return
        if not sheet_path or not os.path.exists(sheet_path):
            QMessageBox.warning(self, "Aviso", "Por favor, selecione uma planilha Excel ou CSV válida.")
            return
        if self.cb_slides.currentIndex() < 0:
            QMessageBox.warning(self, "Aviso", "Por favor, selecione o slide do certificado.")
            return
        if not out_dir:
            QMessageBox.warning(self, "Aviso", "Por favor, selecione uma pasta de destino.")
            return
        if not pattern:
            QMessageBox.warning(self, "Aviso", "Por favor, defina um padrão para o nome do certificado.")
            return
            
        slide_idx = self.cb_slides.itemData(self.cb_slides.currentIndex())
        
        # Cria a pasta de destino se não existir
        if not os.path.exists(out_dir):
            try:
                os.makedirs(out_dir)
                self.txt_log.appendPlainText(f"Pasta de destino criada: {out_dir}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Não foi possível criar a pasta de destino.\n\nDetalhes: {e}")
                return
                
        # 2. Coleta mapeamentos selecionados
        mapping = {}
        for placeholder, cb in self.mapeamentos_widgets.items():
            coluna_selecionada = cb.currentData()
            if coluna_selecionada:
                mapping[placeholder] = coluna_selecionada
                
        if not mapping:
            resp = QMessageBox.question(
                self, "Confirmação",
                "Nenhum campo de substituição foi mapeado. "
                "Isso gerará os certificados exatamente iguais ao modelo.\n\nDeseja continuar mesmo assim?",
                QMessageBox.Yes | QMessageBox.No
            )
            if resp == QMessageBox.No:
                return

        # 3. Determina o formato
        fmt = "PDF" if self.rb_pdf.isChecked() else "PNG"

        # 4. Confirmação do tamanho do lote
        total_certificados = len(self.dados_planilha)
        if total_certificados == 0:
            QMessageBox.warning(self, "Aviso", "A planilha não contém nenhuma linha de dados para gerar certificados.")
            return
            
        # Limpa console de log e barra de progresso
        self.txt_log.clear()
        self.txt_log.appendPlainText(f"Iniciando processo para gerar {total_certificados} certificados...")
        self.bar_progresso.setValue(0)
        self.bar_progresso.setMaximum(total_certificados)
        
        # Desabilita GUI
        self.alternar_estado_gui(True)

        # 5. Inicia a Thread de geração
        self.thread_geracao = QThread()
        self.worker_geracao = GeneratorWorker(
            pptx=pptx_path,
            slide_idx=slide_idx,
            data=self.dados_planilha,
            mapping=mapping,
            fmt=fmt,
            out_dir=out_dir,
            pattern=pattern
        )
        self.worker_geracao.moveToThread(self.thread_geracao)
        
        # Conexões de Sinais
        self.thread_geracao.started.connect(self.worker_geracao.run)
        self.worker_geracao.progress.connect(self.ao_receber_progresso)
        self.worker_geracao.finished.connect(self.ao_concluir_geracao)
        self.worker_geracao.error.connect(self.ao_ocorrer_erro)
        
        # Garbage collection da Thread
        self.worker_geracao.finished.connect(self.thread_geracao.quit)
        self.worker_geracao.finished.connect(self.worker_geracao.deleteLater)
        self.thread_geracao.finished.connect(self.thread_geracao.deleteLater)
        
        self.thread_geracao.start()

    @Slot()
    def cancelar_geracao(self):
        if self.worker_geracao:
            self.txt_log.appendPlainText("Cancelando geração... Aguarde finalizar o certificado atual...")
            self.worker_geracao.cancel()
            self.btn_cancelar.setEnabled(False)

    @Slot(int, int, str)
    def ao_receber_progresso(self, atual, total, mensagem):
        self.bar_progresso.setValue(atual)
        self.txt_log.appendPlainText(mensagem)

    @Slot()
    def ao_concluir_geracao(self):
        self.alternar_estado_gui(False)
        self.txt_log.appendPlainText("\nProcesso concluído com sucesso!")
        
        # Pergunta se o usuário quer abrir a pasta de destino
        out_dir = self.txt_pasta_destino.text()
        resposta = QMessageBox.information(
            self, "Concluído",
            f"Todos os certificados foram gerados na pasta:\n{out_dir}\n\nDeseja abrir a pasta agora?",
            QMessageBox.Open | QMessageBox.Close
        )
        if resposta == QMessageBox.Open:
            os.startfile(out_dir)

    @Slot(str)
    def ao_ocorrer_erro(self, erro_msg):
        self.thread_geracao.quit()
        self.alternar_estado_gui(False)
        self.txt_log.appendPlainText(f"\n[ERRO] Ocorreu um erro no processo: {erro_msg}")
        QMessageBox.critical(
            self, "Erro na Geração",
            f"Ocorreu um erro ao gerar os certificados.\n\nDetalhes: {erro_msg}"
        )
