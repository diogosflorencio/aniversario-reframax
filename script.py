from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QComboBox, QPushButton, QFileDialog,
                             QMessageBox, QScrollArea, QGroupBox, QGridLayout,
                             QSystemTrayIcon, QMenu, QCheckBox, QSpinBox, QTextBrowser,
                             QTabWidget, QColorDialog, QDoubleSpinBox)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QPixmap, QImage, QAction, QIcon, QPainter, QColor, QFont, QKeySequence, QShortcut
from PIL import Image, ImageDraw, ImageFont
from datetime import date
import locale
import pandas as pd
import os
import glob
import sys
import json

import layout_config
import layout_editor
import preferencias
import servidor_licenca
from version import VERSION

PASTA_OUTPUTS = "outputs"
COR_TEXTO_PADRAO = (43, 57, 100)


def obter_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

# Configuração da localização para português
locale.setlocale(locale.LC_TIME, 'pt_BR')

class AniversariantesApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.base_dir = obter_base_dir()
        self.df = None
        self.imagem = None
        self.aniversariantes_info = []
        self.layouts_disponiveis = []
        self.prefs = preferencias.carregar(self.base_dir)
        self.forcar_saida = False
        self.tray_aviso_mostrado = False
        self._ignorar_defaults_layout = False
        self._layout_json = None
        self._ignorar_mes_editor = False
        self._ignorar_mes_checks = False
        self._ignorar_alteracao_editor = False
        self._editor_snapshot_salvo = None
        self._editor_tab_index = None
        self._tab_indice_anterior = 0
        self._bloquear_mudanca_tab = False
        self._layout_id_confirmado = None
        self._bloquear_sync_vis = False
        self._ignorar_sync_layout = False
        self._snapshot_titulo_x = 0.125
        self._snapshot_lista_x = 0.125
        self.initUI()
        self.configurar_bandeja()
        self.aplicar_preferencias_na_ui()
        self.iniciar_timer_agendamentos()
        
    def initUI(self):
        self.app_icon = self.criar_icone_bandeja()
        self.setWindowIcon(self.app_icon)
        self.setWindowTitle(f'Gerador de Aniversariantes v{VERSION}')
        self.setGeometry(100, 100, 1280, 820)

        central = QWidget()
        outer_layout = QVBoxLayout(central)
        outer_layout.setContentsMargins(8, 8, 8, 8)
        outer_layout.setSpacing(8)
        self.setCentralWidget(central)

        header = QHBoxLayout()
        header_icon = QLabel()
        header_icon.setPixmap(self.app_icon.pixmap(32, 32))
        header_icon.setFixedSize(36, 36)
        header_title = QLabel(f"Gerador de Aniversariantes v{VERSION}")
        header_title.setStyleSheet("font-size: 15px; font-weight: bold;")
        header.addWidget(header_icon)
        header.addWidget(header_title)
        header.addStretch()
        outer_layout.addLayout(header)

        self.tab_widget = QTabWidget()
        outer_layout.addWidget(self.tab_widget)

        # -- Aba Principal: configurações + pré-visualização --
        principal_tab = QWidget()
        principal_layout = QHBoxLayout(principal_tab)
        principal_layout.setSpacing(12)

        left_panel = QWidget()
        left_panel.setMaximumWidth(480)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setSpacing(10)

        config_group = QGroupBox("Configurações")
        config_layout = QGridLayout()

        self.excel_combo = QComboBox()
        self.atualizar_lista_excel()
        config_layout.addWidget(QLabel('Planilha:'), 0, 0)
        config_layout.addWidget(self.excel_combo, 0, 1)

        self.modelo_layout_combo = QComboBox()
        self.atualizar_lista_layouts()
        self.modelo_layout_combo.currentIndexChanged.connect(self.on_layout_selecionado)
        config_layout.addWidget(QLabel('Layout:'), 1, 0)
        config_layout.addWidget(self.modelo_layout_combo, 1, 1)

        self.mes_label = QLabel('Mês:')
        self.mes_combo = QComboBox()
        meses = ['Mês Atual'] + [date(2020, m, 1).strftime('%B').capitalize() for m in range(1, 13)]
        self.mes_combo.addItems(meses)
        config_layout.addWidget(self.mes_label, 2, 0)
        config_layout.addWidget(self.mes_combo, 2, 1)

        self.ano_label = QLabel('Ano:')
        self.ano_combo = QComboBox()
        ano_atual = date.today().year
        self.ano_combo.addItems(['Ano Atual'] + [str(a) for a in range(ano_atual - 1, ano_atual + 4)])
        config_layout.addWidget(self.ano_label, 2, 0)
        config_layout.addWidget(self.ano_combo, 2, 1)
        self.ano_label.setVisible(False)
        self.ano_combo.setVisible(False)

        self.modelo_fundo_combo = QComboBox()
        self.atualizar_lista_modelos()
        config_layout.addWidget(QLabel('Modelo:'), 3, 0)
        config_layout.addWidget(self.modelo_fundo_combo, 3, 1)

        self.fonte_lista_combo = QComboBox()
        self.fonte_mes_combo = QComboBox()
        self.atualizar_lista_fontes()
        config_layout.addWidget(QLabel('Fonte nomes:'), 4, 0)
        config_layout.addWidget(self.fonte_lista_combo, 4, 1)
        self.fonte_mes_label = QLabel('Fonte mês:')
        config_layout.addWidget(self.fonte_mes_label, 5, 0)
        config_layout.addWidget(self.fonte_mes_combo, 5, 1)

        self.estilo_combo = QComboBox()
        self.estilo_combo.addItems([
            'Estilo padrão',
            'Linhas coloridas',
            'Dias à esquerda'
        ])
        config_layout.addWidget(QLabel('Estilo:'), 6, 0)
        config_layout.addWidget(self.estilo_combo, 6, 1)

        self._cor_ano = COR_TEXTO_PADRAO
        self._cor_mes = COR_TEXTO_PADRAO
        self._cor_nomes = COR_TEXTO_PADRAO

        self.cor_ano_label = QLabel('Cor ano:')
        self.cor_ano_btn = QPushButton()
        self.cor_ano_btn.setFixedWidth(72)
        self.cor_ano_btn.clicked.connect(lambda: self.escolher_cor('ano'))
        config_layout.addWidget(self.cor_ano_label, 7, 0)
        config_layout.addWidget(self.cor_ano_btn, 7, 1)

        self.cor_mes_label = QLabel('Cor mês:')
        self.cor_mes_btn = QPushButton()
        self.cor_mes_btn.setFixedWidth(72)
        self.cor_mes_btn.clicked.connect(lambda: self.escolher_cor('mes'))
        config_layout.addWidget(self.cor_mes_label, 8, 0)
        config_layout.addWidget(self.cor_mes_btn, 8, 1)

        self.cor_nomes_label = QLabel('Cor nomes:')
        self.cor_nomes_btn = QPushButton()
        self.cor_nomes_btn.setFixedWidth(72)
        self.cor_nomes_btn.clicked.connect(lambda: self.escolher_cor('nomes'))
        config_layout.addWidget(self.cor_nomes_label, 9, 0)
        config_layout.addWidget(self.cor_nomes_btn, 9, 1)

        self._atualizar_botoes_cor()

        self._mostrar_ano = True
        self._mostrar_meses = True
        self.chk_mostrar_ano = QCheckBox("Mostrar ano")
        self.chk_mostrar_ano.setChecked(True)
        self.chk_mostrar_ano.toggled.connect(self._on_visibilidade_ano_changed)
        config_layout.addWidget(self.chk_mostrar_ano, 10, 0, 1, 2)

        self.chk_mostrar_meses = QCheckBox("Mostrar título do mês")
        self.chk_mostrar_meses.setChecked(True)
        self.chk_mostrar_meses.toggled.connect(self._on_visibilidade_meses_changed)
        config_layout.addWidget(self.chk_mostrar_meses, 11, 0, 1, 2)

        config_group.setLayout(config_layout)
        left_layout.addWidget(config_group)

        self.status_salvamento_label = QLabel("")
        self.status_salvamento_label.setWordWrap(True)
        self.status_salvamento_label.setStyleSheet("color: #2e7d32; font-size: 13px;")
        left_layout.addWidget(self.status_salvamento_label)

        self.ultimo_caminho_salvo = None

        buttons_layout = QHBoxLayout()
        self.gerar_btn = QPushButton('Gerar (Ctrl+G)')
        self.gerar_btn.clicked.connect(self.gerar_imagem)
        self.salvar_btn = QPushButton('Salvar (Ctrl+S)')
        self.salvar_btn.clicked.connect(self.salvar_imagem)
        self.imprimir_btn = QPushButton('Imprimir (Ctrl+P)')
        self.imprimir_btn.clicked.connect(self.imprimir_imagem)
        self.abrir_local_btn = QPushButton('Abrir local do arquivo')
        self.abrir_local_btn.clicked.connect(self.abrir_local_arquivo_salvo)
        self.abrir_local_btn.setEnabled(False)
        buttons_layout.addWidget(self.gerar_btn)
        buttons_layout.addWidget(self.salvar_btn)
        buttons_layout.addWidget(self.imprimir_btn)
        left_layout.addLayout(buttons_layout)
        left_layout.addWidget(self.abrir_local_btn)
        left_layout.addStretch()

        self.salvar_btn.setEnabled(False)
        self.imprimir_btn.setEnabled(False)

        principal_layout.addWidget(left_panel)

        preview_group = QGroupBox("Pré-visualização")
        preview_layout = QVBoxLayout(preview_group)
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setMinimumSize(480, 640)
        scroll = QScrollArea()
        scroll.setWidget(self.image_label)
        scroll.setWidgetResizable(True)
        preview_layout.addWidget(scroll)
        principal_layout.addWidget(preview_group, 1)

        self.tab_widget.addTab(principal_tab, "Principal")
        self._criar_aba_editor_layout()

        # -- Aba separada: Preferências + Sobre lado a lado --
        info_tab = QWidget()
        info_layout = QHBoxLayout(info_tab)
        info_layout.setSpacing(12)

        prefs_group = QGroupBox("Preferências")
        prefs_scroll = QScrollArea()
        prefs_scroll.setWidgetResizable(True)
        prefs_content = QWidget()
        prefs_layout = QVBoxLayout(prefs_content)

        salvar_prefs_group = QGroupBox("Preferências de geração")
        salvar_prefs_layout = QGridLayout()
        salvar_prefs_layout.addWidget(QLabel(
            "Salve planilha, mês, layout, modelo, fontes e estilo para reutilizar ao abrir o app ou na geração automática."
        ), 0, 0, 1, 2)

        self.carregar_prefs_check = QCheckBox("Carregar preferências salvas ao iniciar")
        self.carregar_prefs_check.setChecked(self.prefs.get("carregar_ao_iniciar", True))
        salvar_prefs_layout.addWidget(self.carregar_prefs_check, 1, 0, 1, 2)

        self.salvar_prefs_btn = QPushButton("Salvar preferências atuais (Ctrl+Shift+S)")
        self.salvar_prefs_btn.clicked.connect(self.salvar_preferencias_ui)
        salvar_prefs_layout.addWidget(self.salvar_prefs_btn, 2, 0, 1, 2)

        self.prefs_status_label = QLabel("")
        salvar_prefs_layout.addWidget(self.prefs_status_label, 3, 0, 1, 2)
        salvar_prefs_group.setLayout(salvar_prefs_layout)
        prefs_layout.addWidget(salvar_prefs_group)

        lembrete_group = QGroupBox("Lembrete no Windows (criar novo projeto)")
        lembrete_layout = QGridLayout()
        lembrete_layout.addWidget(QLabel(
            "Todo mês, no dia escolhido, o app envia uma notificação na bandeja lembrando de preparar a lista."
        ), 0, 0, 1, 2)

        self.lembrete_ativo_check = QCheckBox("Ativar lembrete mensal")
        lembrete_layout.addWidget(self.lembrete_ativo_check, 1, 0, 1, 2)

        lembrete_layout.addWidget(QLabel("Dia do mês:"), 2, 0)
        self.lembrete_dia_spin = QSpinBox()
        self.lembrete_dia_spin.setRange(1, 31)
        lembrete_layout.addWidget(self.lembrete_dia_spin, 2, 1)

        lembrete_group.setLayout(lembrete_layout)
        prefs_layout.addWidget(lembrete_group)

        auto_group = QGroupBox("Geração e impressão automática")
        auto_layout = QGridLayout()
        auto_layout.addWidget(QLabel(
            "No dia escolhido, o app gera a lista com as preferências salvas e pode enviar direto para a impressora."
        ), 0, 0, 1, 2)

        self.auto_ativo_check = QCheckBox("Ativar geração automática mensal")
        auto_layout.addWidget(self.auto_ativo_check, 1, 0, 1, 2)

        auto_layout.addWidget(QLabel("Dia do mês:"), 2, 0)
        self.auto_dia_spin = QSpinBox()
        self.auto_dia_spin.setRange(1, 31)
        auto_layout.addWidget(self.auto_dia_spin, 2, 1)

        self.auto_imprimir_check = QCheckBox("Imprimir automaticamente após gerar")
        self.auto_imprimir_check.setChecked(False)
        auto_layout.addWidget(self.auto_imprimir_check, 3, 0, 1, 2)

        self.auto_status_label = QLabel("")
        auto_layout.addWidget(self.auto_status_label, 4, 0, 1, 2)
        auto_group.setLayout(auto_layout)
        prefs_layout.addWidget(auto_group)

        bandeja_group = QGroupBox("Bandeja do sistema")
        bandeja_layout = QVBoxLayout()
        bandeja_layout.addWidget(QLabel(
            "Ao fechar a janela (X), o app continua rodando na bandeja do Windows.\n"
            "Clique com o botão direito no ícone da bandeja para abrir ou encerrar de fato."
        ))
        bandeja_group.setLayout(bandeja_layout)
        prefs_layout.addWidget(bandeja_group)

        prefs_layout.addStretch()
        prefs_scroll.setWidget(prefs_content)
        prefs_group_layout = QVBoxLayout(prefs_group)
        prefs_group_layout.addWidget(prefs_scroll)

        about_group = QGroupBox("Sobre")
        about_layout = QVBoxLayout(about_group)
        about_text = QTextBrowser()
        about_text.setOpenExternalLinks(True)
        about_text.setReadOnly(True)
        about_text.setHtml(self._html_sobre())
        about_layout.addWidget(about_text)

        info_layout.addWidget(prefs_group, 1)
        info_layout.addWidget(about_group, 1)

        self.tab_widget.addTab(info_tab, "Preferências e Sobre")

        self._tab_indice_anterior = self.tab_widget.currentIndex()
        self.tab_widget.currentChanged.connect(self._on_tab_mudou)

        self._criar_aviso_editor_pendente()
        self.preencher_campos_preferencias()
        self.configurar_atalhos()
        self.atualizar_controles_tipo_layout()

    def _html_sobre(self):
        return f"""
<div style='padding: 40px; font-family: "Montserrat", sans-serif; color: #ffffff;'>
    <h2 style='color: #ffffff; margin-bottom: 20px; text-align: center; font-size: 32px;'>
        Gerador de Aniversariantes - v{VERSION}
    </h2>

    <h3 style='color: #ffffff; margin-top: 30px; padding: 15px; font-size: 28px;'>
        Tutorial de uso (rápido):
    </h3>
    <ol style='color: #ffffff; line-height: 2; margin-left: 20px; font-size: 15px;'>
        <li style='margin-bottom: 20px;'>
            <b style='font-weight: bold;'>Preparação dos Dados:</b>
            <ul style='margin: 15px 0 15px 25px; padding: 15px;'>
                <li>Coloque os arquivos excel (.xlsx) na pasta "planilhas"</li>
                <li>Coloque imagens de fundo na pasta "modelos" e fontes na pasta "fontes"</li>
                <li>Os arquivos devem ter as colunas "Nome" e "Nascimento"</li>
            </ul>
        </li>
        <li style='margin-bottom: 20px;'>
            <b style='font-weight: bold;'>Gerando a Lista:</b>
            <ul style='margin: 15px 0 15px 25px; padding: 15px;'>
                <li>Selecione o arquivo excel com os dados dos aniversariantes</li>
                <li>Escolha o layout - ele define se o cartaz é <b>mensal</b> ou <b>anual</b></li>
                <li>No modo mensal, escolha o mês (ou use o mês atual)</li>
                <li>No modo anual, escolha o ano exibido no cartaz (todos os meses vêm da planilha)</li>
                <li>Selecione o modelo de fundo, as fontes e o estilo da lista</li>
                <li>Clique em "Gerar imagem" para visualizar</li>
            </ul>
        </li>
        <li style='margin-bottom: 20px;'>
            <b style='font-weight: bold;'>Tipos de layout:</b>
            <ul style='margin: 15px 0 15px 25px; padding: 15px;'>
                <li><b>Mensal</b> (<code>"tipo": "mensal"</code>): um mês por cartaz, com o nome do mês no topo e a lista abaixo</li>
                <li><b>Anual</b> (<code>"tipo": "anual"</code>): um cartaz com todos os meses; use <code>meses/1..12</code> no JSON para posicionar título e lista de cada mês</li>
                <li>Layouts ficam em <code>layouts/nome_do_layout/layout.json</code></li>
                <li>Modelos de fundo e fontes aparecem automaticamente nos combos</li>
            </ul>
        </li>
        <li style='margin-bottom: 20px;'>
            <b style='font-weight: bold;'>Estilos da lista:</b>
            <ul style='margin: 15px 0 15px 25px; padding: 15px;'>
                <li>Estilo padrão: nomes ........ dias</li>
                <li>Linhas coloridas: o fundo de cada nome fica colorido alternadamente</li>
                <li>Dias à esquerda: dias ........ nome</li>
            </ul>
        </li>
        <li style='margin-bottom: 20px;'>
            <b style='font-weight: bold;'>Preferências e automação:</b>
            <ul style='margin: 15px 0 15px 25px; padding: 15px;'>
                <li>Na aba <b>Preferências e Sobre</b>, salve planilha, mês ou ano, layout, modelo, fontes e estilo</li>
                <li>Configure lembrete mensal na bandeja do Windows</li>
                <li>Configure geração (e impressão) automática mensal com as preferências salvas</li>
                <li>Ao fechar a janela, o app fica na bandeja - clique com o botão direito no ícone para abrir ou encerrar</li>
            </ul>
        </li>
        <li style='margin-bottom: 20px;'>
            <b style='font-weight: bold;'>Atalhos:</b>
            <ul style='margin: 15px 0 15px 25px; padding: 15px;'>
                <li>Ctrl+G ou F5 - gerar imagem</li>
                <li>Ctrl+S - salvar</li>
                <li>Ctrl+P - imprimir</li>
                <li>Ctrl+Shift+S - salvar preferências</li>
            </ul>
        </li>
        <li style='margin-bottom: 20px;'>
            <b style='font-weight: bold;'>Finalizando:</b>
            <ul style='margin: 15px 0 15px 25px; padding: 15px;'>
                <li>Use o botão "Salvar" para guardar a imagem</li>
                <li>Use o botão "Imprimir" para enviar direto pra a impressora</li>
            </ul>
        </li>
    </ol>

    <div style='padding: 25px; margin-top: 30px;'>
        <p style='font-weight: bold; font-size: 15px; margin-bottom: 15px;'>Observações:</p>
        <ul style='color: #ffffff; margin-left: 25px; line-height: 1.8; font-size: 15px;'>
            <li>O programa processará automaticamente nomes muito longos</li>
            <li>As imagens geradas ficam na pasta <code>outputs/</code></li>
            <li>Em caso de problemas, entre em contato comigo</li>
        </ul>
    </div>

    <p style='color: #ffffff; font-family: "Consolas", monospace; font-size: 15px; margin-top: 35px; text-align: center; padding: 15px;'>
        Desenvolvido por <a href="https://www.linkedin.com/in/diogosflorencio/" style="color: #7eb8ff; text-decoration: none;">Diogo</a>
    </p>
</div>
"""

    def configurar_atalhos(self):
        QShortcut(QKeySequence("Ctrl+G"), self, self.gerar_imagem)
        QShortcut(QKeySequence("F5"), self, self.gerar_imagem)
        QShortcut(QKeySequence("Ctrl+S"), self, self.salvar_imagem)
        QShortcut(QKeySequence("Ctrl+P"), self, self.imprimir_imagem)
        QShortcut(QKeySequence("Ctrl+Shift+S"), self, self.salvar_preferencias_ui)

    def on_layout_selecionado(self, _index=-1):
        if self._ignorar_sync_layout:
            return
        novo_id = self.modelo_layout_combo.currentData()
        self._layout_id_confirmado = novo_id
        self.aplicar_defaults_do_layout()
        self.atualizar_controles_tipo_layout()
        self.carregar_editor_do_layout(forcar=True)

    def _selecionar_modelo_compativel(self, nome_referencia):
        if not nome_referencia:
            return
        stem = layout_config._stem_arquivo(nome_referencia)
        for i in range(self.modelo_fundo_combo.count()):
            texto = self.modelo_fundo_combo.itemText(i)
            if layout_config._stem_arquivo(texto) == stem:
                self.modelo_fundo_combo.setCurrentIndex(i)
                return
        self._selecionar_combo_por_texto(self.modelo_fundo_combo, nome_referencia)

    def layout_e_mensal(self, config=None):
        if config is None:
            config = self.obter_layout_selecionado()
        return config is None or layout_config.tipo_layout(config) == "mensal"

    def atualizar_controles_tipo_layout(self, _index=-1):
        config = self.obter_layout_selecionado()
        mensal = self.layout_e_mensal(config)
        self.mes_label.setVisible(mensal)
        self.mes_combo.setVisible(mensal)
        self.ano_label.setVisible(not mensal)
        self.ano_combo.setVisible(not mensal)
        self.fonte_mes_label.setText("Fonte mês:" if mensal else "Fonte títulos:")
        self.cor_ano_label.setVisible(not mensal)
        self.cor_ano_btn.setVisible(not mensal)
        self.cor_mes_label.setText("Cor mês:" if mensal else "Cor meses:")
        self.chk_mostrar_ano.setVisible(not mensal)
        if hasattr(self, "ed_mostrar_ano"):
            self.ed_mostrar_ano.setVisible(not mensal)
        self.chk_mostrar_meses.setText("Mostrar título do mês" if mensal else "Mostrar nomes dos meses")
        if hasattr(self, "ed_mostrar_meses"):
            self.ed_mostrar_meses.setText("Mostrar nome do mês" if mensal else "Mostrar nomes dos meses")
        if hasattr(self, "editor_ano_group"):
            self._atualizar_visibilidade_editor()

    def _on_visibilidade_ano_changed(self, valor):
        if self._bloquear_sync_vis:
            return
        self._mostrar_ano = valor
        self._bloquear_sync_vis = True
        try:
            if self.sender() is not self.chk_mostrar_ano:
                self.chk_mostrar_ano.setChecked(valor)
            if hasattr(self, "ed_mostrar_ano") and self.sender() is not self.ed_mostrar_ano:
                self.ed_mostrar_ano.setChecked(valor)
        finally:
            self._bloquear_sync_vis = False
        self._on_editor_campo_alterado("ano_visivel")

    def _on_visibilidade_meses_changed(self, valor):
        if self._bloquear_sync_vis:
            return
        self._mostrar_meses = valor
        self._bloquear_sync_vis = True
        try:
            if self.sender() is not self.chk_mostrar_meses:
                self.chk_mostrar_meses.setChecked(valor)
            if hasattr(self, "ed_mostrar_meses") and self.sender() is not self.ed_mostrar_meses:
                self.ed_mostrar_meses.setChecked(valor)
        finally:
            self._bloquear_sync_vis = False
        self._on_editor_campo_alterado("titulo_visivel")

    def _aplicar_visibilidade_nos_checks(self, mostrar_ano, mostrar_meses):
        self._bloquear_sync_vis = True
        try:
            self._mostrar_ano = mostrar_ano
            self._mostrar_meses = mostrar_meses
            self.chk_mostrar_ano.setChecked(mostrar_ano)
            self.chk_mostrar_meses.setChecked(mostrar_meses)
            if hasattr(self, "ed_mostrar_ano"):
                self.ed_mostrar_ano.setChecked(mostrar_ano)
            if hasattr(self, "ed_mostrar_meses"):
                self.ed_mostrar_meses.setChecked(mostrar_meses)
        finally:
            self._bloquear_sync_vis = False

    def _visibilidade_meses_anual_do_layout(self):
        meses = (self._layout_json or {}).get("meses", {})
        for m in range(1, 13):
            if meses.get(str(m), {}).get("titulo", {}).get("visivel") is False:
                return False
        return True

    def _aplicar_visibilidade_meses_anual(self, visivel):
        if not self._layout_json:
            return
        for m in range(1, 13):
            titulo = self._layout_json.setdefault("meses", {}).setdefault(str(m), {}).setdefault("titulo", {})
            titulo["visivel"] = visivel

    def _criar_aba_editor_layout(self):
        editor_tab = QWidget()
        editor_layout = QHBoxLayout(editor_tab)

        painel = QWidget()
        painel.setMaximumWidth(420)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(painel)
        form = QVBoxLayout(painel)

        intro = QLabel(
            "Ajuste posições e tamanhos do layout atual. "
            "O preview à direita atualiza automaticamente ao alterar qualquer valor. "
            "Salvar grava no layout.json deste modelo."
        )
        intro.setWordWrap(True)
        form.addWidget(intro)

        self.editor_arquivo_label = QLabel("")
        self.editor_arquivo_label.setWordWrap(True)
        self.editor_arquivo_label.setStyleSheet("color: #555; font-size: 12px;")
        form.addWidget(self.editor_arquivo_label)

        botoes = QHBoxLayout()
        self.editor_preview_btn = QPushButton("Preview layout")
        self.editor_preview_btn.clicked.connect(self.preview_layout_editor)
        self.editor_salvar_btn = QPushButton("Salvar no layout.json")
        self.editor_salvar_btn.clicked.connect(self.salvar_layout_editor)
        self.editor_recarregar_btn = QPushButton("Recarregar")
        self.editor_recarregar_btn.clicked.connect(lambda: self.carregar_editor_do_layout(forcar=False))
        botoes.addWidget(self.editor_preview_btn)
        botoes.addWidget(self.editor_salvar_btn)
        botoes.addWidget(self.editor_recarregar_btn)
        form.addLayout(botoes)

        self.editor_ano_group = QGroupBox("Ano (topo)")
        ano_grid = QGridLayout()
        self.ed_ano_y = self._spin_int(0, 2500, 50)
        self.ed_ano_x = self._spin_x_rel(0.5)
        self.ed_ano_fonte = self._spin_int(8, 400, 90)
        self.ed_mostrar_ano = QCheckBox("Mostrar ano")
        self.ed_mostrar_ano.setChecked(True)
        self.ed_mostrar_ano.toggled.connect(self._on_visibilidade_ano_changed)
        ano_grid.addWidget(QLabel("Y (px)"), 0, 0)
        ano_grid.addWidget(self.ed_ano_y, 0, 1)
        ano_grid.addWidget(QLabel("X (0 = esq, 1 = dir)"), 1, 0)
        ano_grid.addWidget(self.ed_ano_x, 1, 1)
        ano_grid.addWidget(QLabel("Tamanho fonte"), 2, 0)
        ano_grid.addWidget(self.ed_ano_fonte, 2, 1)
        ano_grid.addWidget(self.ed_mostrar_ano, 3, 0, 1, 2)
        self.editor_ano_group.setLayout(ano_grid)
        form.addWidget(self.editor_ano_group)

        titulo_group = QGroupBox("Título do mês")
        titulo_grid = QGridLayout()

        self.ed_meses_selecao_group = QWidget()
        meses_sel_layout = QVBoxLayout(self.ed_meses_selecao_group)
        meses_sel_layout.setContentsMargins(0, 0, 0, 0)
        meses_sel_layout.addWidget(QLabel("Aplicar edições em:"))
        meses_sel_layout.addWidget(QLabel(
            "Marque os meses alvo. Cada parâmetro alterado é salvo só neles, sem mudar os demais campos."
        ))
        self.ed_mes_checks = {}
        meses_grid = QGridLayout()
        abrev_meses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
        for m in range(1, 13):
            cb = QCheckBox(abrev_meses[m - 1])
            cb.setToolTip(layout_config.nome_mes(m))
            cb.stateChanged.connect(self._on_meses_selecao_changed)
            self.ed_mes_checks[m] = cb
            meses_grid.addWidget(cb, (m - 1) // 4, (m - 1) % 4)
        meses_sel_layout.addLayout(meses_grid)
        meses_btns = QHBoxLayout()
        self.ed_meses_todos_btn = QPushButton("Todos")
        self.ed_meses_todos_btn.clicked.connect(self._selecionar_todos_meses_editor)
        self.ed_meses_linha_btn = QPushButton("Mesma linha")
        self.ed_meses_linha_btn.clicked.connect(self._selecionar_meses_linha_referencia)
        self.ed_meses_coluna_btn = QPushButton("Mesma coluna")
        self.ed_meses_coluna_btn.clicked.connect(self._selecionar_meses_coluna_referencia)
        for btn in (
            self.ed_meses_todos_btn,
            self.ed_meses_linha_btn,
            self.ed_meses_coluna_btn,
        ):
            meses_btns.addWidget(btn)
        meses_sel_layout.addLayout(meses_btns)
        self.ed_meses_selecao_label = QLabel("")
        self.ed_meses_selecao_label.setStyleSheet("color: #555; font-size: 12px;")
        meses_sel_layout.addWidget(self.ed_meses_selecao_label)
        titulo_grid.addWidget(self.ed_meses_selecao_group, 0, 0, 1, 2)

        self.ed_titulo_y = self._spin_int(0, 2500, 100)
        self.ed_titulo_x = self._spin_x_rel(0.5)
        self.ed_titulo_fonte = self._spin_int(8, 400, 30)
        self.ed_mostrar_meses = QCheckBox("Mostrar nome do mês")
        self.ed_mostrar_meses.setChecked(True)
        self.ed_mostrar_meses.toggled.connect(self._on_visibilidade_meses_changed)
        titulo_grid.addWidget(QLabel("Y título (px)"), 1, 0)
        titulo_grid.addWidget(self.ed_titulo_y, 1, 1)
        titulo_grid.addWidget(QLabel("X título (0-1, desloca todos)"), 2, 0)
        titulo_grid.addWidget(self.ed_titulo_x, 2, 1)
        titulo_grid.addWidget(QLabel("Fonte título"), 3, 0)
        titulo_grid.addWidget(self.ed_titulo_fonte, 3, 1)
        titulo_grid.addWidget(self.ed_mostrar_meses, 4, 0, 1, 2)
        titulo_group.setLayout(titulo_grid)
        form.addWidget(titulo_group)

        self.ed_espacamento_anual_group = QGroupBox("Espaçamento (meses selecionados)")
        esp_grid = QGridLayout()
        self.ed_gap_titulo_lista = QSpinBox()
        self.ed_gap_titulo_lista.setRange(-300, 800)
        self.ed_gap_titulo_lista.setValue(30)
        self.ed_gap_titulo_lista.setToolTip(
            "Distância em pixels entre o título do mês e a lista. "
            "Negativo aproxima/sobe a lista em direção ao título."
        )
        self.ed_passo_espaco = QSpinBox()
        self.ed_passo_espaco.setRange(1, 50)
        self.ed_passo_espaco.setValue(5)
        self.ed_passo_espaco.setToolTip("Quantos pixels cada clique nos botões move.")

        gap_btns = QHBoxLayout()
        self.ed_gap_menos_btn = QPushButton("− Aproximar")
        self.ed_gap_mais_btn = QPushButton("+ Afastar")
        self.ed_gap_menos_btn.setToolTip("Diminui o espaço vertical (lista sobe)")
        self.ed_gap_mais_btn.setToolTip("Aumenta o espaço vertical (lista desce)")
        self.ed_gap_menos_btn.clicked.connect(lambda: self._nudge_gap_vertical(-1))
        self.ed_gap_mais_btn.clicked.connect(lambda: self._nudge_gap_vertical(1))
        gap_btns.addWidget(self.ed_gap_menos_btn)
        gap_btns.addWidget(self.ed_gap_mais_btn)

        mover_btns = QHBoxLayout()
        self.ed_mover_x_menos_btn = QPushButton("← Esquerda")
        self.ed_mover_x_mais_btn = QPushButton("Direita →")
        self.ed_mover_x_menos_btn.setToolTip("Move todos os meses selecionados para a esquerda")
        self.ed_mover_x_mais_btn.setToolTip("Move todos os meses selecionados para a direita")
        self.ed_mover_x_menos_btn.clicked.connect(lambda: self._nudge_mover_tudo_x(-1))
        self.ed_mover_x_mais_btn.clicked.connect(lambda: self._nudge_mover_tudo_x(1))
        mover_btns.addWidget(self.ed_mover_x_menos_btn)
        mover_btns.addWidget(self.ed_mover_x_mais_btn)

        col_btns = QHBoxLayout()
        self.ed_col_junta_btn = QPushButton("← Junta colunas")
        self.ed_col_afasta_btn = QPushButton("Afasta colunas →")
        self.ed_col_junta_btn.setToolTip("Aproxima colunas entre si (2+ meses)")
        self.ed_col_afasta_btn.setToolTip("Afasta colunas entre si (2+ meses)")
        self.ed_col_junta_btn.clicked.connect(lambda: self._nudge_sep_colunas(-1))
        self.ed_col_afasta_btn.clicked.connect(lambda: self._nudge_sep_colunas(1))
        col_btns.addWidget(self.ed_col_junta_btn)
        col_btns.addWidget(self.ed_col_afasta_btn)

        esp_grid.addWidget(QLabel("Espaço título → lista (px)"), 0, 0)
        esp_grid.addWidget(self.ed_gap_titulo_lista, 0, 1)
        esp_grid.addLayout(gap_btns, 0, 2)
        esp_grid.addWidget(QLabel("Passo (px)"), 1, 0)
        esp_grid.addWidget(self.ed_passo_espaco, 1, 1)
        esp_grid.addWidget(QLabel("Mover todos no eixo X:"), 2, 0)
        esp_grid.addLayout(mover_btns, 2, 1, 1, 2)
        esp_grid.addWidget(QLabel("Entre colunas (2+ meses):"), 3, 0)
        esp_grid.addLayout(col_btns, 3, 1, 1, 2)
        esp_grid.addWidget(QLabel(
            "Use os botões para ajustes finos. Com 1 mês, X (0-1) define posição absoluta; "
            "com vários meses, use ← → para deslocar todos igualmente."
        ), 4, 0, 1, 3)
        self.ed_espacamento_anual_group.setLayout(esp_grid)
        form.addWidget(self.ed_espacamento_anual_group)

        lista_group = QGroupBox("Lista de nomes")
        lista_grid = QGridLayout()
        self.ed_lista_y = self._spin_int(0, 2500, 450)
        self.ed_lista_x = self._spin_x_rel(0.5)
        self.ed_lista_espaco = self._spin_int(5, 200, 70)
        self.ed_lista_fonte = self._spin_int(8, 400, 30)
        self.ed_lista_largura = self._spin_int(50, 1200, 600)
        lista_grid.addWidget(QLabel("Início Y (px)"), 0, 0)
        lista_grid.addWidget(self.ed_lista_y, 0, 1)
        lista_grid.addWidget(QLabel("X lista (0-1, desloca todos)"), 1, 0)
        lista_grid.addWidget(self.ed_lista_x, 1, 1)
        lista_grid.addWidget(QLabel("Espaçamento"), 2, 0)
        lista_grid.addWidget(self.ed_lista_espaco, 2, 1)
        lista_grid.addWidget(QLabel("Fonte nomes"), 3, 0)
        lista_grid.addWidget(self.ed_lista_fonte, 3, 1)
        lista_grid.addWidget(QLabel("Largura linha"), 4, 0)
        lista_grid.addWidget(self.ed_lista_largura, 4, 1)
        lista_group.setLayout(lista_grid)
        form.addWidget(lista_group)

        self.editor_status_label = QLabel("")
        self.editor_status_label.setWordWrap(True)
        form.addWidget(self.editor_status_label)
        form.addStretch()

        editor_layout.addWidget(scroll)

        preview_group = QGroupBox("Pré-visualização do layout")
        preview_layout = QVBoxLayout(preview_group)
        self.editor_image_label = QLabel()
        self.editor_image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.editor_image_label.setMinimumSize(480, 640)
        preview_scroll = QScrollArea()
        preview_scroll.setWidget(self.editor_image_label)
        preview_scroll.setWidgetResizable(True)
        preview_layout.addWidget(preview_scroll)
        editor_layout.addWidget(preview_group, 1)

        self._editor_tab_index = self.tab_widget.addTab(editor_tab, "Editor de layout")
        self._editor_preview_timer = QTimer(self)
        self._editor_preview_timer.setSingleShot(True)
        self._editor_preview_timer.setInterval(450)
        self._editor_preview_timer.timeout.connect(self._preview_layout_editor_auto)
        self._conectar_alteracoes_editor()

    def _criar_aviso_editor_pendente(self):
        self.editor_pendente_label = QLabel("")
        self.editor_pendente_label.setStyleSheet("color: #bf360c; font-size: 12px;")
        self.editor_pendente_label.setVisible(False)
        self.statusBar().addPermanentWidget(self.editor_pendente_label)

    def _conectar_alteracoes_editor(self):
        for widget, campo in (
            (self.ed_ano_y, "ano_y"),
            (self.ed_ano_x, "ano_x"),
            (self.ed_ano_fonte, "ano_fonte"),
            (self.ed_titulo_y, "titulo_y"),
            (self.ed_titulo_x, "titulo_x"),
            (self.ed_titulo_fonte, "titulo_fonte"),
            (self.ed_lista_y, "lista_inicio_y"),
            (self.ed_lista_x, "lista_x"),
            (self.ed_lista_espaco, "lista_espacamento"),
            (self.ed_lista_fonte, "lista_fonte_tamanho"),
            (self.ed_lista_largura, "lista_largura_linha"),
            (self.ed_gap_titulo_lista, "gap_titulo_lista"),
        ):
            widget.valueChanged.connect(
                lambda _v, c=campo: self._on_editor_campo_alterado(c)
            )

    def _on_editor_campo_alterado(self, campo):
        if self._ignorar_mes_editor or self._ignorar_alteracao_editor:
            return
        config = self.obter_layout_selecionado()
        if config and self._layout_json:
            if layout_config.tipo_layout(config) == "anual":
                if campo == "titulo_visivel":
                    self._aplicar_visibilidade_meses_anual(self._mostrar_meses)
                elif campo == "ano_visivel":
                    pass
                elif campo in ("titulo_x", "lista_x"):
                    self._aplicar_x_meses_selecionados(campo)
                elif campo == "gap_titulo_lista":
                    self._aplicar_gap_titulo_lista_meses_selecionados()
                elif not campo.startswith("ano_"):
                    self._salvar_campo_meses_selecionados(campo)
            else:
                self._aplicar_campo_layout_mensal(campo)
        self._agendar_preview_editor()
        self._atualizar_aviso_editor_pendente()

    def _agendar_preview_editor(self):
        if hasattr(self, "_editor_preview_timer"):
            self._editor_preview_timer.start()

    def _snapshot_layout_editor(self):
        if not self._layout_json:
            return None
        return json.dumps(
            layout_editor.clonar_config(self._layout_json),
            sort_keys=True,
            ensure_ascii=False,
        )

    def _editor_tem_alteracoes_pendentes(self):
        if not self._layout_json or self._editor_snapshot_salvo is None:
            return False
        return self._snapshot_layout_editor() != self._editor_snapshot_salvo

    def _atualizar_snapshot_editor_salvo(self):
        if not self._layout_json:
            self._editor_snapshot_salvo = None
        else:
            self._editor_snapshot_salvo = json.dumps(
                layout_editor.clonar_config(self._layout_json),
                sort_keys=True,
                ensure_ascii=False,
            )
        self._atualizar_aviso_editor_pendente()

    def _atualizar_aviso_editor_pendente(self):
        if not hasattr(self, "editor_pendente_label"):
            return
        pendente = self._editor_tem_alteracoes_pendentes()
        self.editor_pendente_label.setVisible(pendente)
        if pendente:
            self.editor_pendente_label.setText("Layout não salvo")
        if self._editor_tab_index is not None:
            self.tab_widget.setTabText(
                self._editor_tab_index,
                "Editor de layout *" if pendente else "Editor de layout",
            )

    def _perguntar_salvar_ou_descartar(self, acao):
        resp = QMessageBox.question(
            self,
            "Alterações não salvas",
            f"Há alterações no editor de layout que ainda não foram salvas.\n\n"
            f"Deseja salvar antes de {acao}?",
            QMessageBox.StandardButton.Save
            | QMessageBox.StandardButton.Discard
            | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        if resp == QMessageBox.StandardButton.Save:
            return "save"
        if resp == QMessageBox.StandardButton.Discard:
            return "discard"
        return None

    def _confirmar_descartar_editor(self, acao):
        if not self._editor_tem_alteracoes_pendentes():
            return True
        escolha = self._perguntar_salvar_ou_descartar(acao)
        if escolha is None:
            return False
        if escolha == "save":
            self.salvar_layout_editor()
            return not self._editor_tem_alteracoes_pendentes()
        if escolha == "discard":
            self._recarregar_layout_editor_do_disco()
            return True
        return False

    def _recarregar_layout_editor_do_disco(self):
        config = self.obter_layout_selecionado()
        if not config or not config.get("_path"):
            self._layout_json = None
            self._editor_snapshot_salvo = None
            self._atualizar_aviso_editor_pendente()
            return
        self._layout_json = layout_editor.clonar_config(
            layout_editor.carregar_arquivo(config["_path"])
        )
        self._carregar_valores_editor_da_memoria()

    def _on_tab_mudou(self, novo_indice):
        if self._bloquear_mudanca_tab:
            return
        anterior = getattr(self, "_tab_indice_anterior", 0)
        editor_idx = self._editor_tab_index
        if editor_idx is not None and anterior == editor_idx and novo_indice != editor_idx:
            if not self._confirmar_descartar_editor("mudar de aba"):
                self._bloquear_mudanca_tab = True
                try:
                    self.tab_widget.setCurrentIndex(anterior)
                finally:
                    self._bloquear_mudanca_tab = False
                return
        self._tab_indice_anterior = novo_indice

    def _spin_int(self, minimo, maximo, valor):
        spin = QSpinBox()
        spin.setRange(minimo, maximo)
        spin.setValue(valor)
        return spin

    def _spin_x_rel(self, valor):
        spin = QDoubleSpinBox()
        spin.setRange(0.0, 1.0)
        spin.setDecimals(3)
        spin.setSingleStep(0.005)
        spin.setValue(valor)
        return spin

    def _altura_modelo_atual(self):
        caminho = self.obter_caminho_modelo_selecionado()
        if caminho and os.path.isfile(caminho):
            with Image.open(caminho) as img:
                return img.height
        return 1000

    def _atualizar_visibilidade_editor(self):
        if not hasattr(self, "editor_ano_group"):
            return
        config = self.obter_layout_selecionado()
        anual = config and layout_config.tipo_layout(config) == "anual"
        self.editor_ano_group.setVisible(bool(anual))
        if hasattr(self, "ed_meses_selecao_group"):
            self.ed_meses_selecao_group.setVisible(bool(anual))
        if hasattr(self, "ed_espacamento_anual_group"):
            self.ed_espacamento_anual_group.setVisible(bool(anual))

    def _largura_modelo_imagem(self):
        caminho = self.obter_caminho_modelo_selecionado()
        if caminho and os.path.isfile(caminho):
            with Image.open(caminho) as img:
                return img.width
        return 1000

    def _x_rel_padrao_mes(self, mes_num):
        return 0.125 * ((mes_num - 1) % 4) + 0.125

    def _obter_bloco_mes_anual(self, mes_num):
        return self._layout_json.setdefault("meses", {}).setdefault(str(mes_num), {})

    def _x_rel_titulo_mes(self, mes_num):
        titulo = self._obter_bloco_mes_anual(mes_num).get("titulo", {})
        return layout_editor.x_rel_para_editor(titulo, self._x_rel_padrao_mes(mes_num))

    def _x_rel_lista_mes(self, mes_num):
        mes_cfg = self._obter_bloco_mes_anual(mes_num)
        lista = mes_cfg.get("lista", {})
        return layout_editor.x_rel_para_editor(lista, self._x_rel_titulo_mes(mes_num))

    def _deslocar_x_mes_anual(self, mes_num, delta_x_rel, mover_titulo=True, mover_lista=True):
        mes_cfg = self._obter_bloco_mes_anual(mes_num)
        if mover_titulo:
            titulo = dict(mes_cfg.get("titulo", {}))
            titulo["x_rel"] = round(
                max(0.0, min(1.0, self._x_rel_titulo_mes(mes_num) + delta_x_rel)),
                4,
            )
            titulo.setdefault("centralizado_h", True)
            mes_cfg["titulo"] = titulo
        if mover_lista:
            lista = dict(mes_cfg.get("lista", {}))
            if mover_titulo:
                lista["x_rel"] = mes_cfg["titulo"]["x_rel"]
            else:
                lista["x_rel"] = round(
                    max(0.0, min(1.0, self._x_rel_lista_mes(mes_num) + delta_x_rel)),
                    4,
                )
            mes_cfg["lista"] = lista

    def _aplicar_x_meses_selecionados(self, campo):
        meses = self._meses_editor_selecionados()
        if not meses or len(meses) > 1:
            return
        snap_attr = "_snapshot_titulo_x" if campo == "titulo_x" else "_snapshot_lista_x"
        novo = float(
            self.ed_titulo_x.value() if campo == "titulo_x" else self.ed_lista_x.value()
        )
        setattr(self, snap_attr, novo)
        self._aplicar_campo_mes_anual(meses[0], campo)
        self._reset_snapshots_espaco_editor()

    def _passo_espaco_px(self):
        return int(self.ed_passo_espaco.value())

    def _nudge_gap_vertical(self, direcao):
        if self._ignorar_alteracao_editor or not self._layout_json:
            return
        passo = self._passo_espaco_px() * direcao
        self._ignorar_alteracao_editor = True
        try:
            self.ed_gap_titulo_lista.setValue(self.ed_gap_titulo_lista.value() + passo)
        finally:
            self._ignorar_alteracao_editor = False
        self._aplicar_gap_titulo_lista_meses_selecionados()
        self._agendar_preview_editor()
        self._atualizar_aviso_editor_pendente()

    def _nudge_mover_tudo_x(self, direcao):
        if self._ignorar_alteracao_editor or not self._layout_json:
            return
        meses = self._meses_editor_selecionados()
        if not meses:
            return
        delta_px = self._passo_espaco_px() * direcao
        dx_rel = delta_px / self._largura_modelo_imagem()
        for mes_num in meses:
            self._deslocar_x_mes_anual(mes_num, dx_rel, mover_titulo=True, mover_lista=True)
        if len(meses) == 1:
            self._sincronizar_x_spins_mes(meses[0])
        self._agendar_preview_editor()
        self._atualizar_aviso_editor_pendente()

    def _nudge_sep_colunas(self, direcao):
        if self._ignorar_alteracao_editor or not self._layout_json:
            return
        meses = self._meses_editor_selecionados()
        if len(meses) < 2:
            return
        delta_px = self._passo_espaco_px() * direcao
        dx_rel = delta_px / self._largura_modelo_imagem()
        ordenados = sorted(meses, key=self._x_rel_titulo_mes)
        n = len(ordenados)
        for i, mes_num in enumerate(ordenados):
            shift = (i - (n - 1) / 2) * dx_rel
            self._deslocar_x_mes_anual(mes_num, shift, mover_titulo=True, mover_lista=True)
        self._agendar_preview_editor()
        self._atualizar_aviso_editor_pendente()

    def _sincronizar_x_spins_mes(self, mes_num):
        self._ignorar_alteracao_editor = True
        try:
            self.ed_titulo_x.setValue(self._x_rel_titulo_mes(mes_num))
            self.ed_lista_x.setValue(self._x_rel_lista_mes(mes_num))
        finally:
            self._ignorar_alteracao_editor = False
        self._reset_snapshots_espaco_editor()

    def _atualizar_controles_espaco_editor(self):
        if not hasattr(self, "ed_col_junta_btn"):
            return
        multi = len(self._meses_editor_selecionados()) > 1
        self.ed_titulo_x.setEnabled(not multi)
        self.ed_lista_x.setEnabled(not multi)
        self.ed_col_junta_btn.setEnabled(multi)
        self.ed_col_afasta_btn.setEnabled(multi)
        self.ed_titulo_x.setToolTip(
            "Posição horizontal do título (0 = esq, 1 = dir)."
            + (" Com vários meses selecionados, use 'Mover todos no eixo X'." if multi else "")
        )
        self.ed_lista_x.setToolTip(
            "Posição horizontal da lista."
            + (" Com vários meses selecionados, use 'Mover todos no eixo X'." if multi else "")
        )

    def _aplicar_gap_titulo_lista_meses_selecionados(self):
        if not self._layout_json:
            return
        gap = int(self.ed_gap_titulo_lista.value())
        altura = self._altura_modelo_atual()
        for mes_num in self._meses_editor_selecionados():
            mes_cfg = self._obter_bloco_mes_anual(mes_num)
            titulo = mes_cfg.get("titulo", {})
            y_titulo = layout_editor.y_para_editor(titulo, altura, 95)
            lista = dict(mes_cfg.get("lista", {}))
            lista["inicio_y"] = int(y_titulo + gap)
            mes_cfg["lista"] = lista
        mes_ancora = self._mes_ancora_selecao_editor()
        inicio = self._obter_bloco_mes_anual(mes_ancora).get("lista", {}).get("inicio_y")
        if inicio is not None:
            self._ignorar_alteracao_editor = True
            try:
                self.ed_lista_y.setValue(int(inicio))
            finally:
                self._ignorar_alteracao_editor = False

    def _reset_snapshots_espaco_editor(self):
        self._snapshot_titulo_x = float(self.ed_titulo_x.value())
        self._snapshot_lista_x = float(self.ed_lista_x.value())

    def _meses_editor_selecionados(self):
        if not hasattr(self, "ed_mes_checks"):
            return [1]
        selecionados = [m for m, cb in self.ed_mes_checks.items() if cb.isChecked()]
        return sorted(selecionados) if selecionados else [1]

    def _mes_ancora_selecao_editor(self):
        return min(self._meses_editor_selecionados())

    def _definir_meses_selecionados(self, meses, sincronizar_formulario=True):
        if not hasattr(self, "ed_mes_checks"):
            return
        if not meses:
            meses = [1]
        meses_set = set(meses)
        self._ignorar_mes_checks = True
        try:
            for m, cb in self.ed_mes_checks.items():
                cb.setChecked(m in meses_set)
        finally:
            self._ignorar_mes_checks = False
        self._atualizar_label_meses_selecionados()
        if sincronizar_formulario:
            self._carregar_formulario_do_primeiro_mes_selecionado()
            self._agendar_preview_editor()

    def _carregar_formulario_do_primeiro_mes_selecionado(self):
        if not self._layout_json:
            return
        self._ignorar_alteracao_editor = True
        try:
            self._preencher_editor_mes_anual(self._mes_ancora_selecao_editor())
        finally:
            self._ignorar_alteracao_editor = False

    def _atualizar_label_meses_selecionados(self):
        if not hasattr(self, "ed_meses_selecao_label"):
            return
        meses = self._meses_editor_selecionados()
        if len(meses) == 1:
            self.ed_meses_selecao_label.setText(
                f"1 mês selecionado ({layout_config.nome_mes(meses[0])})"
            )
        else:
            nomes = ", ".join(layout_config.nome_mes(m)[:3] for m in meses)
            self.ed_meses_selecao_label.setText(f"{len(meses)} meses selecionados: {nomes}")

    def _on_meses_selecao_changed(self, _state=-1):
        if self._ignorar_mes_checks or self._ignorar_mes_editor:
            return
        if not any(cb.isChecked() for cb in self.ed_mes_checks.values()):
            sender = self.sender()
            if isinstance(sender, QCheckBox):
                self._ignorar_mes_checks = True
                sender.setChecked(True)
                self._ignorar_mes_checks = False
            return
        self._carregar_formulario_do_primeiro_mes_selecionado()
        self._atualizar_label_meses_selecionados()
        self._atualizar_controles_espaco_editor()
        self._agendar_preview_editor()
        self._atualizar_aviso_editor_pendente()

    def _selecionar_todos_meses_editor(self):
        self._definir_meses_selecionados(list(range(1, 13)))

    def _selecionar_meses_linha_referencia(self):
        ref = self._mes_ancora_selecao_editor()
        linha = (ref - 1) // 4
        self._definir_meses_selecionados([m for m in range(1, 13) if (m - 1) // 4 == linha])

    def _selecionar_meses_coluna_referencia(self):
        ref = self._mes_ancora_selecao_editor()
        coluna = (ref - 1) % 4
        self._definir_meses_selecionados([m for m in range(1, 13) if (m - 1) % 4 == coluna])

    def _aplicar_campo_mes_anual(self, mes_num, campo):
        if not self._layout_json:
            return
        mes_cfg = self._layout_json.setdefault("meses", {}).setdefault(str(mes_num), {})
        titulo = dict(mes_cfg.get("titulo", {}))
        lista = dict(mes_cfg.get("lista", {}))

        if campo == "titulo_y":
            titulo["y_abs"] = int(self.ed_titulo_y.value())
            titulo.pop("y_rel", None)
        elif campo == "titulo_x":
            titulo["x_rel"] = round(float(self.ed_titulo_x.value()), 4)
            titulo.setdefault("centralizado_h", True)
        elif campo == "titulo_fonte":
            titulo["fonte_tamanho"] = int(self.ed_titulo_fonte.value())
        elif campo == "lista_inicio_y":
            lista["inicio_y"] = int(self.ed_lista_y.value())
        elif campo == "lista_x":
            lista["x_rel"] = round(float(self.ed_lista_x.value()), 4)
        elif campo == "lista_espacamento":
            lista["espacamento"] = int(self.ed_lista_espaco.value())
        elif campo == "lista_fonte_tamanho":
            lista["fonte_tamanho"] = int(self.ed_lista_fonte.value())
        elif campo == "lista_largura_linha":
            lista["largura_linha"] = int(self.ed_lista_largura.value())
        elif campo == "titulo_visivel":
            titulo["visivel"] = self._mostrar_meses
        else:
            return

        if campo.startswith("titulo"):
            mes_cfg["titulo"] = titulo
        elif campo.startswith("lista"):
            mes_cfg["lista"] = lista

    def _salvar_campo_meses_selecionados(self, campo):
        for mes_num in self._meses_editor_selecionados():
            self._aplicar_campo_mes_anual(mes_num, campo)

    def _aplicar_campo_layout_mensal(self, campo):
        if not self._layout_json:
            return
        if campo.startswith("titulo") or campo == "titulo_visivel":
            mes = dict(self._layout_json.get("mes", {}))
            if campo == "titulo_y":
                mes["y_abs"] = int(self.ed_titulo_y.value())
                mes.pop("y_rel", None)
            elif campo == "titulo_x":
                mes["x_rel"] = round(float(self.ed_titulo_x.value()), 4)
                mes.setdefault("centralizado_h", True)
            elif campo == "titulo_fonte":
                mes["fonte_tamanho"] = int(self.ed_titulo_fonte.value())
            elif campo == "titulo_visivel":
                mes["visivel"] = self._mostrar_meses
            else:
                return
            self._layout_json["mes"] = mes
        elif campo.startswith("lista"):
            lista = dict(self._layout_json.get("lista", {}))
            if campo == "lista_inicio_y":
                lista["inicio_y"] = int(self.ed_lista_y.value())
            elif campo == "lista_x":
                lista["x_rel"] = round(float(self.ed_lista_x.value()), 4)
            elif campo == "lista_espacamento":
                lista["espacamento"] = int(self.ed_lista_espaco.value())
            elif campo == "lista_fonte_tamanho":
                lista["fonte_tamanho"] = int(self.ed_lista_fonte.value())
            elif campo == "lista_largura_linha":
                lista["largura_linha"] = int(self.ed_lista_largura.value())
            else:
                return
            self._layout_json["lista"] = lista

    def carregar_editor_do_layout(self, _index=-1, forcar=False):
        if not hasattr(self, "ed_titulo_y"):
            return False
        if not forcar and self._editor_tem_alteracoes_pendentes():
            if not self._confirmar_descartar_editor("recarregar o layout"):
                return False
            if not self._editor_tem_alteracoes_pendentes():
                return True

        config = self.obter_layout_selecionado()
        if not config or not config.get("_path"):
            self._layout_json = None
            self._editor_snapshot_salvo = None
            self.editor_arquivo_label.setText("Nenhum layout selecionado.")
            self._atualizar_aviso_editor_pendente()
            return False

        self._layout_json = layout_editor.clonar_config(
            layout_editor.carregar_arquivo(config["_path"])
        )
        self.editor_arquivo_label.setText(config["_path"])
        self._carregar_valores_editor_da_memoria()
        self.editor_status_label.setStyleSheet("")
        self.editor_status_label.setText("Valores carregados do layout.")
        return True

    def _carregar_valores_editor_da_memoria(self):
        self._atualizar_visibilidade_editor()
        altura = self._altura_modelo_atual()
        config = self.obter_layout_selecionado()
        anual = config and layout_config.tipo_layout(config) == "anual"

        self._ignorar_mes_editor = True
        self._ignorar_alteracao_editor = True
        try:
            if anual:
                titulo_ano = self._layout_json.get("titulo_ano", {})
                self.ed_ano_y.setValue(layout_editor.y_para_editor(titulo_ano, altura, 50))
                self.ed_ano_x.setValue(layout_editor.x_rel_para_editor(titulo_ano, 0.5))
                self.ed_ano_fonte.setValue(int(titulo_ano.get("fonte_tamanho", 90)))

                lista_global = self._layout_json.get("lista", {})
                self.ed_lista_espaco.setValue(int(lista_global.get("espacamento", 20)))
                self.ed_lista_fonte.setValue(int(lista_global.get("fonte_tamanho", 18)))
                self.ed_lista_largura.setValue(int(lista_global.get("largura_linha", 260)))

                self._definir_meses_selecionados([1], sincronizar_formulario=False)
                self._preencher_editor_mes_anual(1, altura)
                self._atualizar_label_meses_selecionados()
                self._aplicar_visibilidade_nos_checks(
                    titulo_ano.get("visivel", True),
                    self._visibilidade_meses_anual_do_layout(),
                )
            else:
                mes = self._layout_json.get("mes", {})
                lista = self._layout_json.get("lista", {})
                self.ed_titulo_y.setValue(layout_editor.y_para_editor(mes, altura, 100))
                self.ed_titulo_x.setValue(layout_editor.x_rel_para_editor(mes, 0.5))
                self.ed_titulo_fonte.setValue(int(mes.get("fonte_tamanho", 200)))
                self.ed_lista_y.setValue(int(lista.get("inicio_y", 450)))
                self.ed_lista_x.setValue(layout_editor.x_rel_para_editor(lista, 0.5))
                self.ed_lista_espaco.setValue(int(lista.get("espacamento", 70)))
                self.ed_lista_fonte.setValue(int(lista.get("fonte_tamanho", 30)))
                self.ed_lista_largura.setValue(int(lista.get("largura_linha", 600)))
                self._aplicar_visibilidade_nos_checks(True, mes.get("visivel", True))
        finally:
            self._ignorar_mes_editor = False
            self._ignorar_alteracao_editor = False

        self._atualizar_snapshot_editor_salvo()
        self._agendar_preview_editor()

    def _preencher_editor_mes_anual(self, mes_num, altura=None):
        if altura is None:
            altura = self._altura_modelo_atual()
        mes_cfg = self._layout_json.setdefault("meses", {}).setdefault(str(mes_num), {})
        titulo = mes_cfg.get("titulo", {})
        lista = mes_cfg.get("lista", {})
        lista_global = self._layout_json.get("lista", {})
        self._ignorar_alteracao_editor = True
        try:
            y_titulo = layout_editor.y_para_editor(titulo, altura, 95)
            inicio_y = int(lista.get("inicio_y", y_titulo + 30))
            self.ed_titulo_y.setValue(y_titulo)
            self.ed_titulo_x.setValue(layout_editor.x_rel_para_editor(titulo, self._x_rel_padrao_mes(mes_num)))
            self.ed_titulo_fonte.setValue(int(titulo.get("fonte_tamanho", 26)))
            self.ed_lista_y.setValue(inicio_y)
            self.ed_lista_x.setValue(layout_editor.x_rel_para_editor(lista, self.ed_titulo_x.value()))
            self.ed_lista_espaco.setValue(int(lista.get("espacamento", lista_global.get("espacamento", 20))))
            self.ed_lista_fonte.setValue(int(lista.get("fonte_tamanho", lista_global.get("fonte_tamanho", 18))))
            self.ed_lista_largura.setValue(int(lista.get("largura_linha", lista_global.get("largura_linha", 260))))
            self.ed_gap_titulo_lista.setValue(inicio_y - y_titulo)
        finally:
            self._ignorar_alteracao_editor = False
        self._reset_snapshots_espaco_editor()
        self._atualizar_controles_espaco_editor()

    def _aplicar_ui_ao_layout_json(self):
        if not self._layout_json:
            return
        config_base = self.obter_layout_selecionado()
        if not config_base:
            return

        cor_mes = list(self._cor_mes)
        cor_nomes = list(self._cor_nomes)
        cor_ano = list(self._cor_ano)

        if layout_config.tipo_layout(config_base) == "anual":
            self._layout_json["titulo_ano"] = self._aplicar_fonte_ao_bloco(
                layout_editor.bloco_texto_de_editor(
                    self.ed_ano_y.value(),
                    self.ed_ano_x.value(),
                    self.ed_ano_fonte.value(),
                    cor_ano,
                    visivel=self._mostrar_ano,
                ),
                self.fonte_mes_combo.currentText(),
            )
            lista_global = self._layout_json.setdefault("lista", {})
            lista_global["cor_texto"] = cor_nomes
            self._aplicar_fonte_ao_bloco(lista_global, self.fonte_lista_combo.currentText())
            self._aplicar_visibilidade_meses_anual(self._mostrar_meses)
        else:
            self._layout_json["mes"] = self._aplicar_fonte_ao_bloco(
                layout_editor.bloco_texto_de_editor(
                    self.ed_titulo_y.value(),
                    self.ed_titulo_x.value(),
                    self.ed_titulo_fonte.value(),
                    cor_mes,
                    visivel=self._mostrar_meses,
                ),
                self.fonte_mes_combo.currentText(),
            )
            self._layout_json["lista"] = self._aplicar_fonte_ao_bloco(
                layout_editor.lista_de_editor(
                    self.ed_lista_y.value(),
                    self.ed_lista_x.value(),
                    self.ed_lista_espaco.value(),
                    self.ed_lista_fonte.value(),
                    self.ed_lista_largura.value(),
                    cor_nomes,
                ),
                self.fonte_lista_combo.currentText(),
            )

    def _config_layout_ativo(self):
        base = self.obter_layout_selecionado()
        if not base:
            return None
        mostrar_ano, mostrar_meses = self._mostrar_ano, self._mostrar_meses
        if not self._layout_json:
            self.carregar_editor_do_layout()
            self._mostrar_ano, self._mostrar_meses = mostrar_ano, mostrar_meses
            self._aplicar_visibilidade_nos_checks(mostrar_ano, mostrar_meses)
        if not self._layout_json:
            return base
        cfg = layout_editor.clonar_config(self._layout_json)
        original = self._layout_json
        try:
            self._layout_json = cfg
            self._aplicar_ui_ao_layout_json()
        finally:
            self._layout_json = original
        cfg["_dir"] = base["_dir"]
        cfg["_path"] = base["_path"]
        return cfg

    def _dataframe_exemplo(self, anual=False):
        linhas = []
        if anual:
            for mes in range(1, 13):
                linhas.append({
                    "Nome": f"Exemplo {layout_config.nome_mes(mes)}",
                    "Nascimento": pd.Timestamp(1990, mes, min(mes * 2, 28)),
                    "Dia": min(mes * 2, 28),
                    "Mes": mes,
                })
            return pd.DataFrame(linhas)
        mes = self.obter_mes_selecionado()
        for dia in (3, 12, 21):
            linhas.append({
                "Nome": f"Aniversariante {dia}",
                "Nascimento": pd.Timestamp(1990, mes, dia),
                "Dia": dia,
                "Mes": mes,
            })
        return pd.DataFrame(linhas)

    def _obter_dados_preview(self):
        try:
            excel = self.excel_combo.currentText()
            if excel:
                arquivo = os.path.join(self.base_dir, "planilhas", excel)
                self.df = pd.read_excel(arquivo)
                return self.df.copy()
        except Exception:
            pass
        config = self.obter_layout_selecionado()
        anual = config and layout_config.tipo_layout(config) == "anual"
        return self._dataframe_exemplo(anual=anual)

    def preview_layout_editor(self):
        self._preview_layout_editor_auto(mostrar_status=True)

    def _preview_layout_editor_auto(self, mostrar_status=False):
        if not self._layout_json:
            return
        ok = self._renderizar_cartaz(
            self._config_layout_ativo(),
            silencioso=True,
            salvar_arquivo=False,
            df_override=self._obter_dados_preview(),
            preview_editor=True,
        )
        if mostrar_status:
            if ok:
                self.editor_status_label.setStyleSheet("color: #2e7d32;")
                self.editor_status_label.setText(
                    "Preview atualizado. Ajuste os valores e salve no layout quando estiver ok."
                )
            else:
                self.editor_status_label.setStyleSheet("color: #c62828;")
                self.editor_status_label.setText(
                    "Não foi possível gerar o preview. Verifique modelo e fontes."
                )
        elif ok:
            self._atualizar_aviso_editor_pendente()

    def salvar_layout_editor(self):
        config = self.obter_layout_selecionado()
        if not config or not config.get("_path"):
            QMessageBox.warning(self, "Salvar layout", "Nenhum layout selecionado.")
            return
        self._aplicar_ui_ao_layout_json()
        try:
            existente = layout_editor.carregar_arquivo(config["_path"])
            for chave, valor in self._layout_json.items():
                existente[chave] = valor
            layout_editor.salvar_arquivo(config["_path"], existente)
            self.atualizar_lista_layouts()
            chave = layout_config.layout_chave(config)
            idx = self.modelo_layout_combo.findData(chave)
            if idx >= 0:
                self.modelo_layout_combo.setCurrentIndex(idx)
            self._atualizar_snapshot_editor_salvo()
            self.editor_status_label.setStyleSheet("color: #2e7d32;")
            self.editor_status_label.setText(f"Layout salvo em:\n{config['_path']}")
        except OSError as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível salvar o layout:\n{e}")

    def obter_ano_selecionado(self):
        if self.ano_combo.currentIndex() == 0:
            return date.today().year
        return int(self.ano_combo.currentText())

    def obter_mes_selecionado(self):
        mes_index = self.mes_combo.currentIndex()
        return date.today().month if mes_index == 0 else mes_index

    def caminho_saida(self):
        pasta = os.path.join(self.base_dir, PASTA_OUTPUTS)
        os.makedirs(pasta, exist_ok=True)
        config = self.obter_layout_selecionado()
        if config and layout_config.tipo_layout(config) == "anual":
            nome = f"aniversariantes_{self.obter_ano_selecionado()}.jpg"
        else:
            mes = layout_config.nome_mes(self.obter_mes_selecionado()).lower()
            nome = f"aniversariantes_{mes}.jpg"
        return os.path.join(pasta, nome)

    def atualizar_lista_excel(self):
        caminho_planilhas = os.path.join(self.base_dir, "planilhas")
        arquivos_excel = glob.glob(os.path.join(caminho_planilhas, "*.xlsx"))
        self.excel_combo.clear()
        for arquivo in arquivos_excel:
            self.excel_combo.addItem(os.path.basename(arquivo))

    def _preencher_combo_arquivos(self, combo, arquivos, vazio_msg):
        selecionado = combo.currentText()
        combo.clear()
        if arquivos:
            combo.addItems(arquivos)
            if selecionado in arquivos:
                combo.setCurrentText(selecionado)
        else:
            combo.addItem(vazio_msg)

    def atualizar_lista_modelos(self):
        modelos = layout_config.listar_modelos(self.base_dir)
        self._preencher_combo_arquivos(self.modelo_fundo_combo, modelos, "Nenhum modelo encontrado")

    def atualizar_lista_fontes(self):
        fontes = layout_config.listar_fontes(self.base_dir)
        self._preencher_combo_arquivos(self.fonte_lista_combo, fontes, "Nenhuma fonte encontrada")
        self._preencher_combo_arquivos(self.fonte_mes_combo, fontes, "Nenhuma fonte encontrada")

    def _selecionar_combo_por_texto(self, combo, texto):
        if texto:
            idx = combo.findText(texto)
            if idx >= 0:
                combo.setCurrentIndex(idx)

    def aplicar_defaults_do_layout(self, _index=-1):
        if self._ignorar_defaults_layout:
            return
        config = self.obter_layout_selecionado()
        if not config:
            return
        defaults = layout_config.defaults_layout(config)
        self._selecionar_modelo_compativel(defaults.get("modelo"))
        self._selecionar_combo_por_texto(self.fonte_lista_combo, defaults.get("fonte_lista"))
        self._selecionar_combo_por_texto(self.fonte_mes_combo, defaults.get("fonte_mes"))
        self.aplicar_cores_do_layout(defaults)

    def _rgb_tuple(self, cor):
        if not cor:
            return COR_TEXTO_PADRAO
        return tuple(int(c) for c in cor[:3])

    def _estilo_botao_cor(self, rgb):
        r, g, b = self._rgb_tuple(rgb)
        return f"background-color: rgb({r}, {g}, {b}); border: 1px solid #666; min-height: 22px;"

    def _atualizar_botoes_cor(self):
        self.cor_ano_btn.setStyleSheet(self._estilo_botao_cor(self._cor_ano))
        self.cor_mes_btn.setStyleSheet(self._estilo_botao_cor(self._cor_mes))
        self.cor_nomes_btn.setStyleSheet(self._estilo_botao_cor(self._cor_nomes))

    def aplicar_cores_do_layout(self, defaults=None):
        if defaults is None:
            config = self.obter_layout_selecionado()
            if not config:
                return
            defaults = layout_config.defaults_layout(config)
        self._cor_ano = self._rgb_tuple(defaults.get("cor_ano"))
        self._cor_mes = self._rgb_tuple(defaults.get("cor_mes"))
        self._cor_nomes = self._rgb_tuple(defaults.get("cor_nomes"))
        self._atualizar_botoes_cor()

    def escolher_cor(self, tipo):
        cores = {"ano": self._cor_ano, "mes": self._cor_mes, "nomes": self._cor_nomes}
        titulos = {"ano": "Cor do ano", "mes": "Cor dos meses", "nomes": "Cor dos nomes"}
        atual = cores.get(tipo, COR_TEXTO_PADRAO)
        escolhida = QColorDialog.getColor(QColor(*atual), self, titulos.get(tipo, "Cor do texto"))
        if not escolhida.isValid():
            return
        rgb = (escolhida.red(), escolhida.green(), escolhida.blue())
        if tipo == "ano":
            self._cor_ano = rgb
        elif tipo == "mes":
            self._cor_mes = rgb
        else:
            self._cor_nomes = rgb
        self._atualizar_botoes_cor()

    def obter_cores_da_ui(self):
        return {
            "ano": self._cor_ano,
            "mes": self._cor_mes,
            "nomes": self._cor_nomes,
        }

    def _cor_efetiva(self, tipo, cor_layout):
        return self.obter_cores_da_ui().get(tipo, COR_TEXTO_PADRAO)

    def obter_caminho_modelo_selecionado(self):
        nome = self.modelo_fundo_combo.currentText()
        if not nome or nome.startswith("Nenhum"):
            return None
        return layout_config.caminho_modelo(self.base_dir, nome)

    def obter_caminho_fonte_lista(self):
        nome = self.fonte_lista_combo.currentText()
        if not nome or nome.startswith("Nenhum"):
            return None
        return layout_config.caminho_fonte(self.base_dir, nome)

    def obter_caminho_fonte_mes(self):
        nome = self.fonte_mes_combo.currentText()
        if not nome or nome.startswith("Nenhum"):
            return None
        return layout_config.caminho_fonte(self.base_dir, nome)

    def _fonte_relativa_layout(self, nome_fonte):
        if not nome_fonte or nome_fonte.startswith("Nenhum"):
            return None
        return layout_editor.caminho_relativo_fonte(nome_fonte)

    def _aplicar_fonte_ao_bloco(self, bloco, nome_fonte):
        rel = self._fonte_relativa_layout(nome_fonte)
        if rel:
            bloco["fonte_arquivo"] = rel
        return bloco

    def atualizar_lista_layouts(self):
        self.layouts_disponiveis = layout_config.listar_layouts(self.base_dir)
        self.modelo_layout_combo.clear()
        for config in self.layouts_disponiveis:
            chave = layout_config.layout_chave(config)
            self.modelo_layout_combo.addItem(
                config.get("nome_exibicao", config.get("id", chave)),
                chave,
            )

        if not self.layouts_disponiveis:
            self.modelo_layout_combo.addItem("Nenhum layout encontrado", None)

    def _selecionar_layout_por_chave(self, chave):
        if not chave:
            return
        idx = self.modelo_layout_combo.findData(chave)
        if idx < 0:
            config = layout_config.carregar_layout(self.base_dir, chave)
            if config:
                idx = self.modelo_layout_combo.findData(layout_config.layout_chave(config))
        if idx >= 0:
            self.modelo_layout_combo.setCurrentIndex(idx)

    def obter_layout_selecionado(self):
        chave = self.modelo_layout_combo.currentData()
        if chave:
            return layout_config.carregar_layout(self.base_dir, chave)
        if self.layouts_disponiveis:
            return self.layouts_disponiveis[0]
        return None

    def preencher_campos_preferencias(self):
        lembrete = self.prefs.get("lembrete_projeto", {})
        auto = self.prefs.get("geracao_automatica", {})

        self.lembrete_ativo_check.setChecked(lembrete.get("ativo", False))
        self.lembrete_dia_spin.setValue(lembrete.get("dia_mes", 25))
        self.auto_ativo_check.setChecked(auto.get("ativo", False))
        self.auto_dia_spin.setValue(auto.get("dia_mes", 1))
        self.auto_imprimir_check.setChecked(auto.get("imprimir", False))
        self.carregar_prefs_check.setChecked(self.prefs.get("carregar_ao_iniciar", True))

        ultima = self.prefs.get("ultima_geracao_auto", "")
        self.auto_status_label.setText(
            f"Última geração automática: {ultima}" if ultima else "Nenhuma geração automática registrada ainda."
        )

    def coletar_preferencias_da_ui(self):
        return {
            "excel": self.excel_combo.currentText(),
            "mes_index": self.mes_combo.currentIndex(),
            "ano_index": self.ano_combo.currentIndex(),
            "layout_id": self.modelo_layout_combo.currentData() or "padrao",
            "modelo": self.modelo_fundo_combo.currentText(),
            "fonte_lista": self.fonte_lista_combo.currentText(),
            "fonte_mes": self.fonte_mes_combo.currentText(),
            "estilo": self.estilo_combo.currentText(),
            "cor_ano": list(self._cor_ano),
            "cor_mes": list(self._cor_mes),
            "cor_nomes": list(self._cor_nomes),
            "carregar_ao_iniciar": self.carregar_prefs_check.isChecked(),
            "lembrete_projeto": {
                "ativo": self.lembrete_ativo_check.isChecked(),
                "dia_mes": self.lembrete_dia_spin.value(),
            },
            "geracao_automatica": {
                "ativo": self.auto_ativo_check.isChecked(),
                "dia_mes": self.auto_dia_spin.value(),
                "imprimir": self.auto_imprimir_check.isChecked(),
            },
            "ultimo_lembrete": self.prefs.get("ultimo_lembrete", ""),
            "ultima_geracao_auto": self.prefs.get("ultima_geracao_auto", ""),
        }

    def aplicar_cores_das_preferencias(self, prefs):
        config = self.obter_layout_selecionado()
        layout_cores = layout_config.defaults_cores(config) if config else {}
        self._cor_ano = self._rgb_tuple(prefs.get("cor_ano") or layout_cores.get("cor_ano"))
        self._cor_mes = self._rgb_tuple(prefs.get("cor_mes") or layout_cores.get("cor_mes"))
        self._cor_nomes = self._rgb_tuple(prefs.get("cor_nomes") or layout_cores.get("cor_nomes"))
        self._atualizar_botoes_cor()

    def salvar_preferencias(self, status_ui=False):
        self.prefs = self.coletar_preferencias_da_ui()
        preferencias.salvar(self.base_dir, self.prefs)
        if status_ui:
            self.prefs_status_label.setText("Preferências salvas com sucesso.")
        return self.prefs

    def salvar_preferencias_ui(self):
        self.salvar_preferencias(status_ui=True)

    def aplicar_preferencias_na_ui(self):
        if not self.prefs.get("carregar_ao_iniciar", True):
            return

        self._ignorar_defaults_layout = True
        self._ignorar_sync_layout = True
        try:
            excel = self.prefs.get("excel", "")
            if excel:
                idx = self.excel_combo.findText(excel)
                if idx >= 0:
                    self.excel_combo.setCurrentIndex(idx)

            mes_index = self.prefs.get("mes_index", 0)
            if 0 <= mes_index < self.mes_combo.count():
                self.mes_combo.setCurrentIndex(mes_index)

            ano_index = self.prefs.get("ano_index", 0)
            if 0 <= ano_index < self.ano_combo.count():
                self.ano_combo.setCurrentIndex(ano_index)

            layout_id = self.prefs.get("layout_id")
            self._selecionar_layout_por_chave(layout_id)

            self.aplicar_defaults_do_layout()
            modelo = self.prefs.get("modelo", "")
            if modelo:
                self._selecionar_modelo_compativel(modelo)

            self._selecionar_combo_por_texto(self.fonte_lista_combo, self.prefs.get("fonte_lista", ""))
            self._selecionar_combo_por_texto(self.fonte_mes_combo, self.prefs.get("fonte_mes", ""))

            estilo = self.prefs.get("estilo", "")
            if estilo:
                idx = self.estilo_combo.findText(estilo)
                if idx >= 0:
                    self.estilo_combo.setCurrentIndex(idx)

            self.aplicar_cores_das_preferencias(self.prefs)
        finally:
            self._ignorar_defaults_layout = False
            self._ignorar_sync_layout = False
        self.atualizar_controles_tipo_layout()
        self._layout_id_confirmado = self.modelo_layout_combo.currentData()
        self.carregar_editor_do_layout(forcar=True)

    def criar_icone_bandeja(self):
        tamanho = 64
        pixmap = QPixmap(tamanho, tamanho)
        pixmap.fill(QColor(28, 32, 38))

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        painter.setPen(QColor(55, 62, 72))
        painter.setBrush(QColor(55, 62, 72))
        painter.drawRoundedRect(4, 4, tamanho - 8, tamanho - 8, 10, 10)

        painter.setPen(QColor(86, 195, 116))
        painter.setFont(QFont("Consolas", 22, QFont.Weight.Bold))
        painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, ">_")

        painter.end()
        return QIcon(pixmap)

    def configurar_bandeja(self):
        icone = getattr(self, "app_icon", None) or self.criar_icone_bandeja()

        self.tray_icon = QSystemTrayIcon(icone, self)
        self.tray_icon.setToolTip("Gerador de Aniversariantes")

        menu = QMenu()
        acao_abrir = QAction("Abrir janela", self)
        acao_abrir.triggered.connect(self.mostrar_janela)
        menu.addAction(acao_abrir)

        acao_gerar = QAction("Gerar imagem", self)
        acao_gerar.triggered.connect(lambda: self.gerar_imagem())
        menu.addAction(acao_gerar)

        menu.addSeparator()

        acao_sair = QAction("Encerrar aplicativo", self)
        acao_sair.triggered.connect(self.encerrar_aplicativo)
        menu.addAction(acao_sair)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self.tray_ativado)
        self.tray_icon.show()

    def tray_ativado(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.mostrar_janela()

    def mostrar_janela(self):
        self.showNormal()
        self.activateWindow()
        self.raise_()

    def encerrar_aplicativo(self):
        if not self._confirmar_descartar_editor("encerrar o aplicativo"):
            return
        self.forcar_saida = True
        self.salvar_preferencias()
        self.tray_icon.hide()
        QApplication.quit()

    def closeEvent(self, event):
        if self.forcar_saida:
            self.salvar_preferencias()
            event.accept()
            return

        if not self._confirmar_descartar_editor("fechar a janela"):
            event.ignore()
            return

        event.ignore()
        self.salvar_preferencias()
        self.hide()

        if not self.tray_aviso_mostrado:
            self.tray_icon.showMessage(
                "Gerador de Aniversariantes",
                "O app continua rodando na bandeja. Clique com o botão direito no ícone para abrir ou encerrar.",
                QSystemTrayIcon.MessageIcon.Information,
                4000,
            )
            self.tray_aviso_mostrado = True

    def iniciar_timer_agendamentos(self):
        self.timer_agendamentos = QTimer(self)
        self.timer_agendamentos.timeout.connect(self.verificar_agendamentos)
        self.timer_agendamentos.start(60_000)
        QTimer.singleShot(3000, self.verificar_agendamentos)

    def verificar_agendamentos(self):
        hoje = preferencias.hoje_iso()
        dia = preferencias.dia_mes_hoje()
        prefs = preferencias.carregar(self.base_dir)

        lembrete = prefs.get("lembrete_projeto", {})
        if lembrete.get("ativo") and dia == lembrete.get("dia_mes"):
            if prefs.get("ultimo_lembrete") != hoje:
                self.tray_icon.showMessage(
                    "Lembrete - lista de aniversariantes",
                    "Hoje é dia de preparar/criar a nova lista de aniversariantes do mês.",
                    QSystemTrayIcon.MessageIcon.Information,
                    8000,
                )
                prefs["ultimo_lembrete"] = hoje
                preferencias.salvar(self.base_dir, prefs)
                self.prefs = prefs

        auto = prefs.get("geracao_automatica", {})
        if auto.get("ativo") and dia == auto.get("dia_mes"):
            if prefs.get("ultima_geracao_auto") != hoje:
                self.executar_geracao_automatica(prefs, auto)

    def executar_geracao_automatica(self, prefs, auto_cfg):
        self.aplicar_preferencias_salvas(prefs)
        ok = self.gerar_imagem(silencioso=True)
        hoje = preferencias.hoje_iso()

        if ok and auto_cfg.get("imprimir", False):
            ok_impressao = self.imprimir_imagem(silencioso=True)
            if ok_impressao:
                self.tray_icon.showMessage(
                    "Geração automática concluída",
                    "Lista gerada e enviada para a impressora com as preferências salvas.",
                    QSystemTrayIcon.MessageIcon.Information,
                    8000,
                )
            else:
                self.tray_icon.showMessage(
                    "Geração automática - erro na impressão",
                    "A imagem foi gerada, mas houve erro ao imprimir. Abra o app para tentar manualmente.",
                    QSystemTrayIcon.MessageIcon.Warning,
                    8000,
                )
        elif ok:
            self.tray_icon.showMessage(
                "Geração automática concluída",
                "Lista gerada e salva com as preferências salvas.",
                QSystemTrayIcon.MessageIcon.Information,
                8000,
            )
        else:
            self.tray_icon.showMessage(
                "Geração automática - falhou",
                "Não foi possível gerar a lista. Abra o app e verifique planilha e configurações.",
                QSystemTrayIcon.MessageIcon.Critical,
                8000,
            )

        prefs["ultima_geracao_auto"] = hoje
        preferencias.salvar(self.base_dir, prefs)
        self.prefs = prefs
        self.auto_status_label.setText(f"Última geração automática: {hoje}")

    def aplicar_preferencias_salvas(self, prefs):
        self.atualizar_lista_modelos()
        self.atualizar_lista_fontes()

        self._ignorar_defaults_layout = True
        self._ignorar_sync_layout = True
        try:
            excel = prefs.get("excel", "")
            if excel:
                idx = self.excel_combo.findText(excel)
                if idx >= 0:
                    self.excel_combo.setCurrentIndex(idx)

            mes_index = prefs.get("mes_index", 0)
            if 0 <= mes_index < self.mes_combo.count():
                self.mes_combo.setCurrentIndex(mes_index)

            ano_index = prefs.get("ano_index", 0)
            if 0 <= ano_index < self.ano_combo.count():
                self.ano_combo.setCurrentIndex(ano_index)

            layout_id = prefs.get("layout_id")
            self._selecionar_layout_por_chave(layout_id)

            self.aplicar_defaults_do_layout()
            modelo = prefs.get("modelo", "")
            if modelo:
                self._selecionar_modelo_compativel(modelo)

            self._selecionar_combo_por_texto(self.fonte_lista_combo, prefs.get("fonte_lista", ""))
            self._selecionar_combo_por_texto(self.fonte_mes_combo, prefs.get("fonte_mes", ""))

            estilo = prefs.get("estilo", "")
            if estilo:
                idx = self.estilo_combo.findText(estilo)
                if idx >= 0:
                    self.estilo_combo.setCurrentIndex(idx)

            self.aplicar_cores_das_preferencias(prefs)
        finally:
            self._ignorar_defaults_layout = False
            self._ignorar_sync_layout = False
        self.atualizar_controles_tipo_layout()
        self._layout_id_confirmado = self.modelo_layout_combo.currentData()

    def notificar(self, titulo, mensagem, silencioso, icone=QMessageBox.Icon.Information):
        if silencioso and self.tray_icon:
            tray_icon = QSystemTrayIcon.MessageIcon.Information
            if icone == QMessageBox.Icon.Warning:
                tray_icon = QSystemTrayIcon.MessageIcon.Warning
            elif icone == QMessageBox.Icon.Critical:
                tray_icon = QSystemTrayIcon.MessageIcon.Critical
            self.tray_icon.showMessage(titulo, mensagem, tray_icon, 6000)
        else:
            if icone == QMessageBox.Icon.Warning:
                QMessageBox.warning(self, titulo, mensagem)
            elif icone == QMessageBox.Icon.Critical:
                QMessageBox.critical(self, titulo, mensagem)
            else:
                QMessageBox.information(self, titulo, mensagem)

    def carregar_dados_excel(self, silencioso=False):
        try:
            arquivo = os.path.join(self.base_dir, "planilhas", self.excel_combo.currentText())
            self.df = pd.read_excel(arquivo)
            return True
        except Exception as e:
            self.notificar('Erro', f'Erro ao carregar arquivo Excel: {str(e)}', silencioso, QMessageBox.Icon.Critical)
            return False

    def preparar_dados(self):
        self.df['Nascimento'] = pd.to_datetime(self.df['Nascimento'])
        self.df['Dia'] = self.df['Nascimento'].dt.day
        self.df['Mes'] = self.df['Nascimento'].dt.month

    def filtrar_aniversariantes(self):
        self.preparar_dados()
        mes = self.obter_mes_selecionado()
        aniversariantes = self.df[self.df['Mes'] == mes].copy()
        return aniversariantes.sort_values(by='Dia')

    def agrupar_aniversariantes_por_mes(self):
        self.preparar_dados()
        grupos = {}
        for mes in range(1, 13):
            grupo = self.df[self.df['Mes'] == mes].copy()
            grupos[mes] = grupo.sort_values(by='Dia')
        return grupos

    def total_aniversariantes_grupos(self, grupos):
        return sum(len(g) for g in grupos.values())

    def processar_nome(self, nome_completo):
        """
        formata o nome do aniversariante com as regras de abreviacao
        1. se tiver 5 palavras, abrevia a segunda (ou terceira se a segunda for conectivo)
        2 se tiver 6 palavras: abrevia duas palavras seguindo a mesma logica - eu removi isso pq ficou desnecessario
        3 se o nome for muito grande (mais que 25 caracteres): abrevia seguindo a mesma logica
        """
        palavras = nome_completo.split()
        conectivos = ['DA', 'DE', 'DO', 'DOS', 'DAS']
        
        # se o nome for pequeno e tiver menos que 5 palavras, retorna normal
        if len(nome_completo) <= 25 and len(palavras) < 5:
            return nome_completo
            
        # decide quantas palavras abreviar baseado no tamanho do nome
        num_abreviacoes = 0
        if len(palavras) >= 6:
            num_abreviacoes = 1 # removi a opção de apreviação de 2 palavras pq coloquei uma abreviação de nomes grandes
        elif len(palavras) == 5 or len(nome_completo) > 25:
            num_abreviacoes = 1
            
        if num_abreviacoes > 0:
            pos_atual = 1  # começa da segunda palavra
            abreviacoes_feitas = 0
            
            while abreviacoes_feitas < num_abreviacoes and pos_atual < len(palavras) - 1:
                # pula se for conectivo
                if palavras[pos_atual].upper() in conectivos:
                    pos_atual += 1
                    continue
                    
                # abrevia a palavra atual
                palavras[pos_atual] = f"{palavras[pos_atual][0]}."
                abreviacoes_feitas += 1
                pos_atual += 1
        
        return ' '.join(palavras)

    def desenhar_forma(self, desenho, forma_cfg, x, y, largura, altura):
        if not forma_cfg or forma_cfg.get("tipo") in (None, "nenhuma"):
            return

        padding_h = forma_cfg.get("padding_h", 15)
        padding_v = forma_cfg.get("padding_v", 20)
        cor = tuple(forma_cfg.get("cor", [230, 236, 255]))
        borda_offset = forma_cfg.get("borda_offset", 1)
        escurecer = forma_cfg.get("escurecer_borda", 30)

        box = [
            (x - padding_h, y - padding_v),
            (x + largura + padding_h, y + altura + padding_v),
        ]
        tipo = forma_cfg.get("tipo", "retangulo")

        if borda_offset:
            cor_borda = tuple(max(0, c - escurecer) for c in cor)
            borda_box = [
                (box[0][0] - borda_offset, box[0][1] - borda_offset),
                (box[1][0] + borda_offset, box[1][1] + borda_offset),
            ]
            if tipo == "elipse":
                desenho.ellipse(borda_box, fill=cor_borda)
            else:
                desenho.rectangle(borda_box, fill=cor_borda)

        if tipo == "elipse":
            desenho.ellipse(box, fill=cor)
        else:
            desenho.rectangle(box, fill=cor)

    def gerar_imagem(self, silencioso=False):
        print("Iniciando geração da imagem...")
        self.atualizar_lista_modelos()
        self.atualizar_lista_fontes()

        if not self.carregar_dados_excel(silencioso=silencioso):
            print("Erro ao carregar dados do Excel")
            return False

        config_layout = self._config_layout_ativo()
        if not config_layout:
            self.notificar('Erro', 'Nenhum layout encontrado na pasta "layouts/".', silencioso, QMessageBox.Icon.Critical)
            return False

        return self._renderizar_cartaz(
            config_layout,
            silencioso=silencioso,
            salvar_arquivo=True,
            df_override=None,
            preview_editor=False,
        )

    def _renderizar_cartaz(self, config_layout, silencioso=False, salvar_arquivo=True, df_override=None, preview_editor=False):
        caminho_imagem = self.obter_caminho_modelo_selecionado()
        caminho_fonte_lista = self.obter_caminho_fonte_lista()
        caminho_fonte_mes = self.obter_caminho_fonte_mes()

        if not caminho_imagem:
            self.notificar('Erro', 'Nenhum modelo de fundo encontrado na pasta "modelos/".', silencioso, QMessageBox.Icon.Critical)
            return False
        if not caminho_fonte_lista or not caminho_fonte_mes:
            self.notificar('Erro', 'Nenhuma fonte encontrada na pasta "fontes/".', silencioso, QMessageBox.Icon.Critical)
            return False

        df_anterior = self.df
        if df_override is not None:
            self.df = df_override

        try:
            print(f"Tentando carregar imagem de: {caminho_imagem}")
            if not os.path.isfile(caminho_imagem):
                self.notificar('Erro', f'Imagem de fundo não encontrada:\n{caminho_imagem}', silencioso, QMessageBox.Icon.Critical)
                return False

            self.imagem = Image.open(caminho_imagem)
            cfg_mes = layout_config.posicao_mes(config_layout, self.imagem.height)
            eh_anual = layout_config.tipo_layout(config_layout) == "anual"

            self._fonte_lista_ui = caminho_fonte_lista
            self._fonte_titulo_ui = caminho_fonte_mes

            desenho = ImageDraw.Draw(self.imagem)
            estilo = self.estilo_combo.currentText()

            if eh_anual:
                self.preparar_dados()
                grupos = self.agrupar_aniversariantes_por_mes()
                total = self.total_aniversariantes_grupos(grupos)
                if total == 0 and not preview_editor:
                    self.notificar(
                        'Aviso',
                        'Nenhum aniversariante encontrado na planilha.',
                        silencioso,
                        QMessageBox.Icon.Warning,
                    )
                    return False

                ano = self.obter_ano_selecionado()
                self.desenhar_titulo_ano(desenho, config_layout, ano)

                for mes_num in range(1, 13):
                    grupo = grupos[mes_num]
                    self.desenhar_titulo_mes(desenho, config_layout, mes_num)
                    if len(grupo) > 0:
                        self.desenhar_lista(desenho, grupo, config_layout, estilo, mes_num)
            else:
                self.preparar_dados()
                aniversariantes = self.filtrar_aniversariantes()
                if len(aniversariantes) == 0 and not preview_editor:
                    self.notificar(
                        'Aviso',
                        'Nenhum aniversariante foi encontrado para o mês selecionado (caso haja de fato, falar com Diogo)!',
                        silencioso,
                        QMessageBox.Icon.Warning,
                    )
                    return False

                texto_mes = layout_config.nome_mes(self.obter_mes_selecionado())
                if cfg_mes.get("visivel", True):
                    self.desenhar_titulo_central(desenho, texto_mes, cfg_mes, "mes")
                self.desenhar_lista(desenho, aniversariantes, config_layout, estilo)

            if not silencioso or preview_editor:
                self.atualizar_preview()

            if salvar_arquivo:
                caminho_saida = self.caminho_saida()
                self.imagem.save(caminho_saida)
                if not silencioso:
                    self.informar_salvamento(caminho_saida)
                self.salvar_btn.setEnabled(True)
                self.imprimir_btn.setEnabled(True)

            return True

        except Exception as e:
            import traceback
            print(f"Erro detalhado: {traceback.format_exc()}")
            self.notificar('Erro', f'Erro ao gerar imagem (se persistir, falar com diogo): {str(e)}', silencioso, QMessageBox.Icon.Critical)
            return False
        finally:
            self.df = df_anterior

    def _x_centralizado_coluna(self, cfg_lista, largura_bloco):
        return layout_config.anchor_x_bloco(cfg_lista, largura_bloco, self.imagem.width)

    def _carregar_fonte(self, cfg, ui_fallback):
        for caminho in (ui_fallback, cfg.get("fonte_arquivo")):
            if caminho and os.path.isfile(caminho):
                return ImageFont.truetype(caminho, cfg["fonte_tamanho"])
        padrao = layout_config.caminho_fonte(self.base_dir, "OpenSans-Regular.ttf")
        if os.path.isfile(padrao):
            return ImageFont.truetype(padrao, cfg["fonte_tamanho"])
        raise FileNotFoundError(
            f'Nenhuma fonte válida encontrada. Coloque arquivos .ttf/.otf na pasta "fontes/".'
        )

    def desenhar_titulo_central(self, desenho, texto, cfg, tipo_cor="mes"):
        fonte = self._carregar_fonte(cfg, self._fonte_titulo_ui)
        text_length = desenho.textlength(texto, fonte)
        x = layout_config.calcular_x_texto(cfg, text_length, self.imagem.width)
        cor = self._cor_efetiva(tipo_cor, cfg["cor"])
        desenho.text((x, cfg["y"]), texto, cor, font=fonte)

    def desenhar_titulo_ano(self, desenho, config_layout, ano):
        if not config_layout.get("titulo_ano"):
            return
        cfg = layout_config.config_titulo_ano(config_layout, self.imagem.height)
        if not cfg.get("visivel", True):
            return
        self.desenhar_titulo_central(desenho, str(ano), cfg, "ano")

    def desenhar_titulo_mes(self, desenho, config_layout, mes_numero):
        cfg = layout_config.config_titulo_mes(config_layout, mes_numero, self.imagem.height)
        if not cfg.get("visivel", True):
            return
        if cfg.get("y") is None and cfg.get("x_rel") is None and cfg.get("x_abs") is None:
            return
        texto = layout_config.nome_mes(mes_numero)
        self.desenhar_titulo_central(desenho, texto, cfg, "mes")

    def desenhar_lista(self, desenho, aniversariantes, config_layout, estilo, mes_numero=None):
        if estilo == 'Linhas coloridas':
            self.desenhar_estilo_colorido(desenho, aniversariantes, config_layout, mes_numero)
        elif estilo == 'Dias à esquerda':
            self.desenhar_estilo_dia_esquerda(desenho, aniversariantes, config_layout, mes_numero)
        else:
            self.desenhar_estilo_padrao(desenho, aniversariantes, config_layout, mes_numero)

    def desenhar_estilo_colorido(self, desenho, aniversariantes, config_layout, mes_numero=None):
        cfg_lista = layout_config.config_lista(config_layout, mes_numero)
        fonte_texto = self._carregar_fonte(cfg_lista, self._fonte_lista_ui)
        cfg_estilo = layout_config.config_estilo_layout(config_layout, "colorido")
        forma_base = layout_config.forma_padrao(config_layout) or {}
        desloc = cfg_lista.get("deslocamento_y", 0)

        pos_init = cfg_lista["inicio_y"]
        espacamento = cfg_estilo.get("espacamento", cfg_lista["espacamento"])
        cores = cfg_estilo.get("cores_fundo", [[230, 236, 255]])
        text_color = self._cor_efetiva("nomes", cfg_lista["cor_texto"])
        self.aniversariantes_info = []

        for i, (_, row) in enumerate(aniversariantes.iterrows()):
            nome = self.processar_nome(row['Nome'])
            dia = str(row['Dia'])
            self.aniversariantes_info.append(f"Nome: {nome} - Dia: {dia}")

            slot = layout_config.slot_para_indice(config_layout, i, pos_init, mes_numero)
            y = slot["y"] + desloc
            texto = f"{nome} - Dia {dia}"
            text_length = desenho.textlength(texto, fonte_texto)

            if slot.get("nome_x_rel") is not None:
                x = self.imagem.width * slot["nome_x_rel"]
            else:
                x = self._x_centralizado_coluna(cfg_lista, text_length)

            forma_cfg = dict(forma_base)
            if slot.get("forma"):
                forma_cfg.update(slot["forma"])
            forma_cfg["cor"] = forma_cfg.get("cor") or cores[i % len(cores)]

            altura_texto = cfg_lista["fonte_tamanho"]
            self.desenhar_forma(desenho, forma_cfg, x, y - 10, text_length, altura_texto)
            desenho.text((x, y - 10), texto, fill=text_color, font=fonte_texto)
            pos_init += espacamento

    def desenhar_estilo_dia_esquerda(self, desenho, aniversariantes, config_layout, mes_numero=None):
        cfg_lista = layout_config.config_lista(config_layout, mes_numero)
        fonte_texto = self._carregar_fonte(cfg_lista, self._fonte_lista_ui)
        cfg_estilo = layout_config.config_estilo_layout(config_layout, "dia_esquerda")
        desloc = cfg_lista.get("deslocamento_y", 0)

        pos_init = cfg_lista["inicio_y"]
        espacamento = cfg_estilo.get("espacamento", cfg_lista["espacamento"])
        text_color = self._cor_efetiva("nomes", cfg_lista["cor_texto"])
        largura_linha = cfg_lista["largura_linha"]
        dia_x_rel_padrao = cfg_estilo.get("dia_x_rel", 0.28)
        self.aniversariantes_info = []

        for i, (_, row) in enumerate(aniversariantes.iterrows()):
            nome = self.processar_nome(row['Nome'])
            dia = str(row['Dia'])
            self.aniversariantes_info.append(f"Nome: {nome} - Dia: {dia}")

            slot = layout_config.slot_para_indice(config_layout, i, pos_init, mes_numero)
            y = slot["y"] + desloc

            dia_x_rel = slot.get("dia_x_rel") if slot.get("dia_x_rel") is not None else dia_x_rel_padrao
            if cfg_lista.get("x_rel") is not None or cfg_lista.get("x_abs") is not None:
                x_dia = self._x_centralizado_coluna(cfg_lista, largura_linha)
            else:
                x_dia = self.imagem.width * dia_x_rel

            dia_width = desenho.textlength(dia, fonte_texto)
            nome_width = desenho.textlength(nome, fonte_texto)
            ponto_width = desenho.textlength('.', fonte_texto)

            largura_pontos = largura_linha - (nome_width + dia_width)
            num_pontos = max(int(largura_pontos // ponto_width), 2)

            forma_cfg = slot.get("forma")
            if forma_cfg:
                altura_texto = cfg_lista["fonte_tamanho"]
                self.desenhar_forma(desenho, forma_cfg, x_dia - dia_width / 2, y - 15, largura_linha, altura_texto)

            desenho.text((x_dia, y - 15), dia, fill=text_color, font=fonte_texto)

            pontos = '.' * num_pontos
            x_pontos = x_dia + dia_width / 2 + 20
            desenho.text((x_pontos, y - 15), pontos, fill=text_color, font=fonte_texto)

            x_nome = x_pontos + (num_pontos * ponto_width) + ponto_width
            if slot.get("nome_x_rel") is not None:
                x_nome = self.imagem.width * slot["nome_x_rel"]
            desenho.text((x_nome, y - 15), nome, fill=text_color, font=fonte_texto)

            pos_init += espacamento

    def desenhar_estilo_padrao(self, desenho, aniversariantes, config_layout, mes_numero=None):
        cfg_lista = layout_config.config_lista(config_layout, mes_numero)
        fonte_texto = self._carregar_fonte(cfg_lista, self._fonte_lista_ui)
        desloc = cfg_lista.get("deslocamento_y", 0)

        pos_init = cfg_lista["inicio_y"]
        espacamento = cfg_lista["espacamento"]
        text_color = self._cor_efetiva("nomes", cfg_lista["cor_texto"])
        largura_linha = cfg_lista["largura_linha"]
        self.aniversariantes_info = []

        for i, (_, row) in enumerate(aniversariantes.iterrows()):
            nome = self.processar_nome(row['Nome'])
            dia = str(row['Dia'])
            self.aniversariantes_info.append(f"Nome: {nome} - Dia: {dia}")

            slot = layout_config.slot_para_indice(config_layout, i, pos_init, mes_numero)
            y = slot["y"] + desloc

            largura_nome = desenho.textlength(nome, fonte_texto)
            largura_dia = desenho.textlength(dia, fonte_texto)
            largura_ponto = desenho.textlength('.', fonte_texto)

            largura_pontos = largura_linha - (largura_nome + largura_dia)
            num_pontos = max(int(largura_pontos // largura_ponto), 2)
            bloco_largura = largura_nome + (num_pontos * largura_ponto) + largura_dia

            if slot.get("nome_x_rel") is not None and slot.get("dia_x_rel") is not None:
                x_nome = self.imagem.width * slot["nome_x_rel"]
                x_dia = self.imagem.width * slot["dia_x_rel"]
                x_pontos = x_nome + largura_nome + largura_ponto
            else:
                x = self._x_centralizado_coluna(cfg_lista, bloco_largura)
                x_nome = x
                x_pontos = x_nome + largura_nome + largura_ponto
                x_dia = x_pontos + (num_pontos * largura_ponto) + largura_ponto

            forma_cfg = slot.get("forma")
            if forma_cfg:
                altura_texto = cfg_lista["fonte_tamanho"]
                self.desenhar_forma(desenho, forma_cfg, min(x_nome, x_dia), y, bloco_largura, altura_texto)

            desenho.text((x_nome, y), nome, fill=text_color, font=fonte_texto)
            desenho.text((x_pontos, y), '.' * num_pontos, fill=text_color, font=fonte_texto)
            desenho.text((x_dia, y), dia, fill=text_color, font=fonte_texto)

            pos_init += espacamento

    def atualizar_preview(self):
        if not self.imagem:
            return
        try:
            img_qt = self.imagem.convert('RGB')
            data = img_qt.tobytes('raw', 'RGB')
            qim = QImage(data, img_qt.size[0], img_qt.size[1], img_qt.size[0] * 3, QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(qim)

            for label in (self.image_label, getattr(self, "editor_image_label", None)):
                if not label:
                    continue
                scaled = pixmap.scaled(
                    label.size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
                label.setPixmap(scaled)
        except Exception as e:
            import traceback
            print(f"Erro ao atualizar preview: {traceback.format_exc()}")
            if not getattr(self, "_preview_silencioso", False):
                QMessageBox.critical(self, 'Erro', f'Erro ao atualizar preview (se persistir, falar com Diogo): {str(e)}')

    def informar_salvamento(self, caminho_saida):
        self.ultimo_caminho_salvo = caminho_saida
        self.abrir_local_btn.setEnabled(os.path.isfile(caminho_saida))
        self.status_salvamento_label.setStyleSheet("color: #2e7d32; font-size: 13px;")
        self.status_salvamento_label.setText(
            f"Imagem salva com sucesso.\nSalvo em: {caminho_saida}"
        )

    def abrir_local_arquivo_salvo(self):
        caminho = self.ultimo_caminho_salvo or self.caminho_saida()
        if not os.path.isfile(caminho):
            QMessageBox.information(
                self,
                "Arquivo não encontrado",
                "Gere ou salve a imagem antes de abrir o local do arquivo.",
            )
            return
        try:
            import subprocess
            subprocess.run(["explorer", "/select,", os.path.normpath(caminho)], check=False)
        except Exception as e:
            QMessageBox.warning(self, "Erro", f"Não foi possível abrir o local do arquivo: {e}")

    def salvar_imagem(self):
        if self.imagem:
            try:
                caminho_saida = self.caminho_saida()
                self.imagem.save(caminho_saida)
                self.informar_salvamento(caminho_saida)
            except Exception as e:
                self.status_salvamento_label.setStyleSheet("color: #c62828; font-size: 13px;")
                self.status_salvamento_label.setText(f"Erro ao salvar: {str(e)}")

    def imprimir_imagem(self, silencioso=False):
        if not self.imagem:
            return False
        try:
            import win32print
            import win32ui
            from PIL import ImageWin

            PHYSICALWIDTH = 110
            PHYSICALHEIGHT = 111

            hDC = win32ui.CreateDC()
            impressora_padrao = win32print.GetDefaultPrinter()
            hDC.CreatePrinterDC(impressora_padrao)
            printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)

            hDC.StartDoc("aniversariantes_output.jpg")
            hDC.StartPage()

            dib = ImageWin.Dib(self.imagem)
            dib.draw(hDC.GetHandleOutput(), (0, 0, printer_size[0], printer_size[1]))

            hDC.EndPage()
            hDC.EndDoc()
            hDC.DeleteDC()

            self.notificar('Deu certo!!', 'Imagem enviada pra impressora!', silencioso)
            return True
        except Exception as e:
            self.notificar('Erro', f'Erro ao imprimir (se persistir, falar com Diogo): {str(e)}', silencioso, QMessageBox.Icon.Critical)
            return False

def main():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    if not QSystemTrayIcon.isSystemTrayAvailable():
        QMessageBox.critical(None, "Bandeja indisponível", "A bandeja do sistema não está disponível neste computador.")
        sys.exit(1)

    ativo, mensagem = servidor_licenca.verificar_servidor_ativo()
    if not ativo:
        QMessageBox.critical(None, "Servidor", mensagem)
        sys.exit(1)

    ex = AniversariantesApp()
    ex.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()