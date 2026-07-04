from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QComboBox, QPushButton, QFileDialog,
                             QMessageBox, QScrollArea, QGroupBox, QGridLayout,
                             QSystemTrayIcon, QMenu, QCheckBox, QSpinBox, QTextBrowser,
                             QTabWidget)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QPixmap, QImage, QAction, QIcon, QPainter, QColor, QFont, QKeySequence, QShortcut
from PIL import Image, ImageDraw, ImageFont
from datetime import date
import locale
import pandas as pd
import os
import glob
import sys

import layout_config
import preferencias
import servidor_licenca
from version import VERSION

PASTA_OUTPUTS = "outputs"

# Configuração da localização para português
locale.setlocale(locale.LC_TIME, 'pt_BR')

class AniversariantesApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.df = None
        self.imagem = None
        self.aniversariantes_info = []
        self.layouts_disponiveis = []
        self.prefs = preferencias.carregar(self.base_dir)
        self.forcar_saida = False
        self.tray_aviso_mostrado = False
        self._ignorar_defaults_layout = False
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
        self.modelo_layout_combo.currentIndexChanged.connect(self.aplicar_defaults_do_layout)
        self.modelo_layout_combo.currentIndexChanged.connect(self.atualizar_controles_tipo_layout)
        config_layout.addWidget(QLabel('Layout:'), 1, 0)
        config_layout.addWidget(self.modelo_layout_combo, 1, 1)

        self.tipo_layout_label = QLabel("")
        self.tipo_layout_label.setWordWrap(True)
        self.tipo_layout_label.setStyleSheet("color: #555; font-size: 12px; font-style: italic;")
        config_layout.addWidget(self.tipo_layout_label, 2, 0, 1, 2)

        self.mes_label = QLabel('Mês:')
        self.mes_combo = QComboBox()
        meses = ['Mês Atual'] + [date(2020, m, 1).strftime('%B').capitalize() for m in range(1, 13)]
        self.mes_combo.addItems(meses)
        config_layout.addWidget(self.mes_label, 3, 0)
        config_layout.addWidget(self.mes_combo, 3, 1)

        self.ano_label = QLabel('Ano:')
        self.ano_combo = QComboBox()
        ano_atual = date.today().year
        self.ano_combo.addItems(['Ano Atual'] + [str(a) for a in range(ano_atual - 1, ano_atual + 4)])
        config_layout.addWidget(self.ano_label, 3, 0)
        config_layout.addWidget(self.ano_combo, 3, 1)
        self.ano_label.setVisible(False)
        self.ano_combo.setVisible(False)

        self.modelo_fundo_combo = QComboBox()
        self.atualizar_lista_modelos()
        config_layout.addWidget(QLabel('Modelo:'), 4, 0)
        config_layout.addWidget(self.modelo_fundo_combo, 4, 1)

        self.fonte_lista_combo = QComboBox()
        self.fonte_mes_combo = QComboBox()
        self.atualizar_lista_fontes()
        config_layout.addWidget(QLabel('Fonte nomes:'), 5, 0)
        config_layout.addWidget(self.fonte_lista_combo, 5, 1)
        self.fonte_mes_label = QLabel('Fonte mês:')
        config_layout.addWidget(self.fonte_mes_label, 6, 0)
        config_layout.addWidget(self.fonte_mes_combo, 6, 1)

        self.estilo_combo = QComboBox()
        self.estilo_combo.addItems([
            'Estilo padrão',
            'Linhas coloridas',
            'Dias à esquerda'
        ])
        config_layout.addWidget(QLabel('Estilo:'), 7, 0)
        config_layout.addWidget(self.estilo_combo, 7, 1)

        config_group.setLayout(config_layout)
        left_layout.addWidget(config_group)

        self.status_salvamento_label = QLabel("")
        self.status_salvamento_label.setWordWrap(True)
        self.status_salvamento_label.setStyleSheet("color: #2e7d32; font-size: 13px;")
        left_layout.addWidget(self.status_salvamento_label)

        buttons_layout = QHBoxLayout()
        self.gerar_btn = QPushButton('Gerar (Ctrl+G)')
        self.gerar_btn.clicked.connect(self.gerar_imagem)
        self.salvar_btn = QPushButton('Salvar (Ctrl+S)')
        self.salvar_btn.clicked.connect(self.salvar_imagem)
        self.imprimir_btn = QPushButton('Imprimir (Ctrl+P)')
        self.imprimir_btn.clicked.connect(self.imprimir_imagem)
        buttons_layout.addWidget(self.gerar_btn)
        buttons_layout.addWidget(self.salvar_btn)
        buttons_layout.addWidget(self.imprimir_btn)
        left_layout.addLayout(buttons_layout)
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
        if mensal:
            modo = "Modo mensal - escolha o mês e gere a lista desse período."
        else:
            modo = "Modo anual - todos os aniversariantes da planilha, agrupados por mês no cartaz."
        descricao = (config or {}).get("descricao", "")
        self.tipo_layout_label.setText(f"{modo} {descricao}".strip())

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
        self._selecionar_combo_por_texto(self.modelo_fundo_combo, defaults.get("modelo"))
        self._selecionar_combo_por_texto(self.fonte_lista_combo, defaults.get("fonte_lista"))
        self._selecionar_combo_por_texto(self.fonte_mes_combo, defaults.get("fonte_mes"))

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

    def atualizar_lista_layouts(self):
        self.layouts_disponiveis = layout_config.listar_layouts(self.base_dir)
        self.modelo_layout_combo.clear()
        for config in self.layouts_disponiveis:
            self.modelo_layout_combo.addItem(config.get("nome_exibicao", config.get("id", "Layout")), config.get("id"))

        if not self.layouts_disponiveis:
            self.modelo_layout_combo.addItem("Nenhum layout encontrado", None)

    def obter_layout_selecionado(self):
        layout_id = self.modelo_layout_combo.currentData()
        if layout_id:
            return layout_config.carregar_layout(self.base_dir, layout_id)
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
        self.auto_imprimir_check.setChecked(auto.get("imprimir", True))
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
            if layout_id:
                idx = self.modelo_layout_combo.findData(layout_id)
                if idx >= 0:
                    self.modelo_layout_combo.setCurrentIndex(idx)

            modelo = self.prefs.get("modelo", "")
            if modelo:
                self._selecionar_combo_por_texto(self.modelo_fundo_combo, modelo)
            else:
                self.aplicar_defaults_do_layout()

            self._selecionar_combo_por_texto(self.fonte_lista_combo, self.prefs.get("fonte_lista", ""))
            self._selecionar_combo_por_texto(self.fonte_mes_combo, self.prefs.get("fonte_mes", ""))

            estilo = self.prefs.get("estilo", "")
            if estilo:
                idx = self.estilo_combo.findText(estilo)
                if idx >= 0:
                    self.estilo_combo.setCurrentIndex(idx)
        finally:
            self._ignorar_defaults_layout = False
        self.atualizar_controles_tipo_layout()

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
        self.forcar_saida = True
        self.salvar_preferencias()
        self.tray_icon.hide()
        QApplication.quit()

    def closeEvent(self, event):
        if self.forcar_saida:
            self.salvar_preferencias()
            event.accept()
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

        if ok and auto_cfg.get("imprimir", True):
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
        if layout_id:
            idx = self.modelo_layout_combo.findData(layout_id)
            if idx >= 0:
                self.modelo_layout_combo.setCurrentIndex(idx)

        self._selecionar_combo_por_texto(self.modelo_fundo_combo, prefs.get("modelo", ""))
        self._selecionar_combo_por_texto(self.fonte_lista_combo, prefs.get("fonte_lista", ""))
        self._selecionar_combo_por_texto(self.fonte_mes_combo, prefs.get("fonte_mes", ""))

        estilo = prefs.get("estilo", "")
        if estilo:
            idx = self.estilo_combo.findText(estilo)
            if idx >= 0:
                self.estilo_combo.setCurrentIndex(idx)

        self.atualizar_controles_tipo_layout()

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

    def _x_centralizado_coluna(self, cfg_lista, largura_bloco):
        x_rel = cfg_lista.get("x_rel")
        if x_rel is not None:
            return self.imagem.width * x_rel - largura_bloco / 2
        return (self.imagem.width - largura_bloco) / 2

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

        config_layout = self.obter_layout_selecionado()
        if not config_layout:
            self.notificar('Erro', 'Nenhum layout encontrado na pasta "layouts/".', silencioso, QMessageBox.Icon.Critical)
            return False

        caminho_imagem = self.obter_caminho_modelo_selecionado()
        caminho_fonte_lista = self.obter_caminho_fonte_lista()
        caminho_fonte_mes = self.obter_caminho_fonte_mes()

        if not caminho_imagem:
            self.notificar('Erro', 'Nenhum modelo de fundo encontrado na pasta "modelos/".', silencioso, QMessageBox.Icon.Critical)
            return False
        if not caminho_fonte_lista or not caminho_fonte_mes:
            self.notificar('Erro', 'Nenhuma fonte encontrada na pasta "fontes/".', silencioso, QMessageBox.Icon.Critical)
            return False

        try:
            print(f"Tentando carregar imagem de: {caminho_imagem}")
            if not os.path.isfile(caminho_imagem):
                self.notificar('Erro', f'Imagem de fundo não encontrada:\n{caminho_imagem}', silencioso, QMessageBox.Icon.Critical)
                return False

            self.imagem = Image.open(caminho_imagem)
            print("Imagem base carregada com sucesso")

            cfg_lista = layout_config.config_lista(config_layout)
            cfg_mes = layout_config.posicao_mes(config_layout, self.imagem.height)
            eh_anual = layout_config.tipo_layout(config_layout) == "anual"

            fonte_texto = ImageFont.truetype(caminho_fonte_lista, cfg_lista["fonte_tamanho"])
            fonte_mes = ImageFont.truetype(caminho_fonte_mes, cfg_mes["fonte_tamanho"])
            print("Fontes carregadas com sucesso")

            self.imagem = Image.open(caminho_imagem)
            desenho = ImageDraw.Draw(self.imagem)
            estilo = self.estilo_combo.currentText()

            if eh_anual:
                grupos = self.agrupar_aniversariantes_por_mes()
                total = self.total_aniversariantes_grupos(grupos)
                print(f"Encontrados {total} aniversariantes no cartaz anual")
                if total == 0:
                    self.notificar(
                        'Aviso',
                        'Nenhum aniversariante encontrado na planilha.',
                        silencioso,
                        QMessageBox.Icon.Warning,
                    )
                    return False

                ano = self.obter_ano_selecionado()
                self.desenhar_titulo_ano(desenho, config_layout, fonte_mes, ano)

                for mes_num in range(1, 13):
                    grupo = grupos[mes_num]
                    if len(grupo) == 0:
                        continue
                    self.desenhar_titulo_mes(desenho, config_layout, fonte_mes, mes_num)
                    self.desenhar_lista(desenho, grupo, fonte_texto, config_layout, estilo, mes_num)
            else:
                aniversariantes = self.filtrar_aniversariantes()
                print(f"Encontrados {len(aniversariantes)} aniversariantes")
                if len(aniversariantes) == 0:
                    self.notificar(
                        'Aviso',
                        'Nenhum aniversariante foi encontrado para o mês selecionado (caso haja de fato, falar com Diogo)!',
                        silencioso,
                        QMessageBox.Icon.Warning,
                    )
                    return False

                texto_mes = layout_config.nome_mes(self.obter_mes_selecionado())
                print(f"Desenhando mês: {texto_mes}")
                self.desenhar_titulo_central(desenho, texto_mes, cfg_mes, fonte_mes)
                self.desenhar_lista(desenho, aniversariantes, fonte_texto, config_layout, estilo)

            print("Atualizando preview...")
            if not silencioso:
                self.atualizar_preview()

            caminho_saida = self.caminho_saida()
            self.imagem.save(caminho_saida)
            if not silencioso:
                self.informar_salvamento(caminho_saida)

            self.salvar_btn.setEnabled(True)
            self.imprimir_btn.setEnabled(True)
            print("Geração de imagem concluída com sucesso!")
            return True

        except Exception as e:
            import traceback
            print(f"Erro detalhado: {traceback.format_exc()}")
            self.notificar('Erro', f'Erro ao gerar imagem (se persistir, falar com diogo): {str(e)}', silencioso, QMessageBox.Icon.Critical)
            return False

    def desenhar_titulo_central(self, desenho, texto, cfg, fonte):
        text_length = desenho.textlength(texto, fonte)
        x = (self.imagem.width - text_length) / 2 if cfg["centralizado_h"] else cfg.get("x_abs", 0)
        desenho.text((x, cfg["y"]), texto, cfg["cor"], font=fonte)

    def desenhar_titulo_ano(self, desenho, config_layout, fonte_mes, ano):
        if not config_layout.get("titulo_ano"):
            return
        cfg = layout_config.config_titulo_ano(config_layout, self.imagem.height)
        fonte = ImageFont.truetype(self.obter_caminho_fonte_mes(), cfg["fonte_tamanho"])
        self.desenhar_titulo_central(desenho, str(ano), cfg, fonte)

    def desenhar_titulo_mes(self, desenho, config_layout, fonte_mes, mes_numero):
        cfg = layout_config.config_titulo_mes(config_layout, mes_numero, self.imagem.height)
        if not cfg.get("visivel", True):
            return
        if cfg.get("y") is None and cfg.get("x_rel") is None:
            return
        texto = layout_config.nome_mes(mes_numero)
        fonte = ImageFont.truetype(self.obter_caminho_fonte_mes(), cfg["fonte_tamanho"])
        text_length = desenho.textlength(texto, fonte)
        if cfg.get("x_rel") is not None:
            x = self.imagem.width * cfg["x_rel"] - text_length / 2
        elif cfg["centralizado_h"]:
            x = (self.imagem.width - text_length) / 2
        else:
            x = cfg.get("x_abs", 0)
        desenho.text((x, cfg["y"]), texto, cfg["cor"], font=fonte)

    def desenhar_lista(self, desenho, aniversariantes, fonte_texto, config_layout, estilo, mes_numero=None):
        if estilo == 'Linhas coloridas':
            self.desenhar_estilo_colorido(desenho, aniversariantes, fonte_texto, config_layout, mes_numero)
        elif estilo == 'Dias à esquerda':
            self.desenhar_estilo_dia_esquerda(desenho, aniversariantes, fonte_texto, config_layout, mes_numero)
        else:
            self.desenhar_estilo_padrao(desenho, aniversariantes, fonte_texto, config_layout, mes_numero)

    def desenhar_estilo_colorido(self, desenho, aniversariantes, fonte_texto, config_layout, mes_numero=None):
        cfg_lista = layout_config.config_lista(config_layout, mes_numero)
        cfg_estilo = layout_config.config_estilo_layout(config_layout, "colorido")
        forma_base = layout_config.forma_padrao(config_layout) or {}

        pos_init = cfg_lista["inicio_y"]
        espacamento = cfg_estilo.get("espacamento", cfg_lista["espacamento"])
        cores = cfg_estilo.get("cores_fundo", [[230, 236, 255]])
        text_color = cfg_lista["cor_texto"]
        self.aniversariantes_info = []

        for i, (_, row) in enumerate(aniversariantes.iterrows()):
            nome = self.processar_nome(row['Nome'])
            dia = str(row['Dia'])
            self.aniversariantes_info.append(f"Nome: {nome} - Dia: {dia}")

            slot = layout_config.slot_para_indice(config_layout, i, pos_init, mes_numero)
            y = slot["y"]
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

    def desenhar_estilo_dia_esquerda(self, desenho, aniversariantes, fonte_texto, config_layout, mes_numero=None):
        cfg_lista = layout_config.config_lista(config_layout, mes_numero)
        cfg_estilo = layout_config.config_estilo_layout(config_layout, "dia_esquerda")

        pos_init = cfg_lista["inicio_y"]
        espacamento = cfg_estilo.get("espacamento", cfg_lista["espacamento"])
        text_color = cfg_lista["cor_texto"]
        largura_linha = cfg_lista["largura_linha"]
        dia_x_rel_padrao = cfg_estilo.get("dia_x_rel", 0.28)
        self.aniversariantes_info = []

        for i, (_, row) in enumerate(aniversariantes.iterrows()):
            nome = self.processar_nome(row['Nome'])
            dia = str(row['Dia'])
            self.aniversariantes_info.append(f"Nome: {nome} - Dia: {dia}")

            slot = layout_config.slot_para_indice(config_layout, i, pos_init, mes_numero)
            y = slot["y"]

            dia_x_rel = slot.get("dia_x_rel") if slot.get("dia_x_rel") is not None else dia_x_rel_padrao
            if cfg_lista.get("x_rel") is not None:
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
                bloco_largura = largura_linha
                self.desenhar_forma(desenho, forma_cfg, x_dia - dia_width / 2, y - 15, bloco_largura, altura_texto)

            desenho.text((x_dia, y - 15), dia, fill=text_color, font=fonte_texto)

            pontos = '.' * num_pontos
            x_pontos = x_dia + dia_width / 2 + 20
            desenho.text((x_pontos, y - 15), pontos, fill=text_color, font=fonte_texto)

            x_nome = x_pontos + (num_pontos * ponto_width) + ponto_width
            if slot.get("nome_x_rel") is not None:
                x_nome = self.imagem.width * slot["nome_x_rel"]
            desenho.text((x_nome, y - 15), nome, fill=text_color, font=fonte_texto)

            pos_init += espacamento

    def desenhar_estilo_padrao(self, desenho, aniversariantes, fonte_texto, config_layout, mes_numero=None):
        cfg_lista = layout_config.config_lista(config_layout, mes_numero)

        pos_init = cfg_lista["inicio_y"]
        espacamento = cfg_lista["espacamento"]
        text_color = cfg_lista["cor_texto"]
        largura_linha = cfg_lista["largura_linha"]
        self.aniversariantes_info = []

        for i, (_, row) in enumerate(aniversariantes.iterrows()):
            nome = self.processar_nome(row['Nome'])
            dia = str(row['Dia'])
            self.aniversariantes_info.append(f"Nome: {nome} - Dia: {dia}")

            slot = layout_config.slot_para_indice(config_layout, i, pos_init, mes_numero)
            y = slot["y"]

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
        if self.imagem:
            try:
                print("Convertendo imagem para preview...")
                # Converte PIL Image para QPixmap
                img_qt = self.imagem.convert('RGB')
                data = img_qt.tobytes('raw', 'RGB')
                qim = QImage(data, img_qt.size[0], img_qt.size[1], img_qt.size[0] * 3, QImage.Format.Format_RGB888)
                pixmap = QPixmap.fromImage(qim)
                
                # Calcula o tamanho para manter a proporção
                label_size = self.image_label.size()
                scaled_pixmap = pixmap.scaled(label_size, 
                                            Qt.AspectRatioMode.KeepAspectRatio,
                                            Qt.TransformationMode.SmoothTransformation)
                
                self.image_label.setPixmap(scaled_pixmap)
                print("Preview atualizado com sucesso!")
            except Exception as e:
                import traceback
                print(f"Erro ao atualizar preview: {traceback.format_exc()}")
                QMessageBox.critical(self, 'Erro', f'Erro ao atualizar preview (se persistir, falar com Diogo): {str(e)}')

    def informar_salvamento(self, caminho_saida):
        self.status_salvamento_label.setStyleSheet("color: #2e7d32; font-size: 13px;")
        self.status_salvamento_label.setText(
            f"Imagem salva com sucesso.\nSalvo em: {caminho_saida}"
        )

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