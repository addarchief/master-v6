import os
import sys
import socket
import pyodbc
import winreg
import unicodedata
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

class SQLServerConnector(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Conector SQL Server - Exportador de Datos")
        self.resize(950, 680)
        self.setMinimumSize(880, 640)
        
        self.connection = None
        self.connected_instance = ""
        self.selected_database = ""
        self.logo_pixmap = self.load_logo()

        self.central = QWidget()
        self.setCentralWidget(self.central)

        self.central_layout = QVBoxLayout(self.central)
        self.central_layout.setContentsMargins(0, 0, 0, 0)
        self.central_layout.setSpacing(0)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area_widget = QWidget()
        self.scroll_area_layout = QVBoxLayout(self.scroll_area_widget)
        self.scroll_area_layout.setContentsMargins(32, 32, 32, 32)
        self.scroll_area_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self.scroll_area.setWidget(self.scroll_area_widget)
        self.central_layout.addWidget(self.scroll_area)

        self.stacked = QStackedWidget()
        self.scroll_area_layout.addWidget(self.stacked, alignment=Qt.AlignmentFlag.AlignHCenter)

        self.setup_connection_page()
        self.setup_database_page()
        self.setup_export_page()

        self.apply_stylesheet()
        self.apply_window_icon()
        self.stacked.setCurrentWidget(self.connection_page)

    def get_resource_path(self, relative_path):
        """Obtiene la ruta absoluta al recurso, funciona para desarrollo y para PyInstaller"""
        try:
            # PyInstaller crea una carpeta temporal y almacena la ruta en _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        
        return os.path.join(base_path, relative_path)
    
    def load_logo(self):
        """Carga el logo desde diferentes ubicaciones posibles"""
        # Lista de posibles ubicaciones del logo (en orden de prioridad)
        possible_paths = [
            self.get_resource_path("resources/logo.png"),  # Para PyInstaller
            self.get_resource_path("logo.png"),            # Para PyInstaller (raíz)
            os.path.join(os.path.dirname(__file__), "resources", "logo.png"),  # Desarrollo
            os.path.join(os.path.dirname(__file__), "logo.png"),  # Desarrollo (raíz)
            "resources/logo.png",  # Ruta relativa
            "logo.png",            # Ruta relativa directa
        ]
        
        for logo_path in possible_paths:
            if os.path.exists(logo_path):
                pixmap = QPixmap(logo_path)
                if not pixmap.isNull():
                    print(f"✅ Logo cargado desde: {logo_path}")
                    return pixmap
        
        # Si no se encuentra el logo, mostrar mensaje (solo en desarrollo)
        print("⚠️ No se pudo cargar el logo. Verifica que el archivo 'logo.png' exista en la carpeta 'resources'.")
        return None

    # ----------- Configuración de páginas ----------- #
    def setup_connection_page(self):
        self.connection_page = QWidget()
        layout = QVBoxLayout(self.connection_page)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(24)

        # Logo centrado
        logo_label = self._create_logo_label(80)
        if logo_label:
            logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(logo_label, alignment=Qt.AlignmentFlag.AlignCenter)

        title = QLabel("Conector SQL Server")
        title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        subtitle = QLabel("Conecta y exporta datos a archivos TXT")
        subtitle.setObjectName("subtitleLabel")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(subtitle)

        card = QGroupBox("Credenciales de conexión")
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(16)

        search_layout = QVBoxLayout()
        search_label = QLabel("🔍 Buscar Instancias SQL Server")
        search_label.setObjectName("sectionLabel")
        search_layout.addWidget(search_label)

        self.search_btn = QPushButton("Buscar Instancias Disponibles")
        self.search_btn.clicked.connect(self.search_instances)
        self.search_btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.search_btn.setMinimumWidth(240)
        search_layout.addWidget(self.search_btn, alignment=Qt.AlignmentFlag.AlignLeft)

        instance_layout = QHBoxLayout()
        instance_label = QLabel("Instancia:")
        self.instance_combo = QComboBox()
        self.instance_combo.setEditable(True)
        instance_layout.addWidget(instance_label)
        instance_layout.addWidget(self.instance_combo)
        search_layout.addLayout(instance_layout)

        card_layout.addLayout(search_layout)

        self.windows_check = QCheckBox("Usar autenticación de Windows (Recomendado)")
        self.windows_check.setChecked(True)
        self.windows_check.stateChanged.connect(self.toggle_credentials)
        card_layout.addWidget(self.windows_check)

        credentials_form = QFormLayout()
        credentials_form.setSpacing(12)
        self.user_entry = QLineEdit()
        self.pass_entry = QLineEdit()
        self.pass_entry.setEchoMode(QLineEdit.EchoMode.Password)
        credentials_form.addRow("Usuario:", self.user_entry)
        credentials_form.addRow("Contraseña:", self.pass_entry)
        card_layout.addLayout(credentials_form)

        self.connect_btn = QPushButton("Conectar al Servidor")
        self.connect_btn.setEnabled(False)
        self.connect_btn.clicked.connect(self.connect_to_server)
        card_layout.addWidget(self.connect_btn, alignment=Qt.AlignmentFlag.AlignHCenter)

        self.status_label = QLabel("Listo para conectar")
        self.status_label.setObjectName("infoLabel")
        card_layout.addWidget(self.status_label)

        help_label = QLabel("💡 Consejo: Usa autenticación de Windows para una conexión más segura")
        help_label.setObjectName("hintLabel")
        card_layout.addWidget(help_label)

        layout.addWidget(card)
        layout.addStretch()
        self.stacked.addWidget(self.connection_page)
        self.toggle_credentials()

    def setup_database_page(self):
        self.database_page = QWidget()
        layout = QVBoxLayout(self.database_page)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(24)

        # Logo centrado
        logo_label = self._create_logo_label(80)
        if logo_label:
            logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(logo_label, alignment=Qt.AlignmentFlag.AlignCenter)

        title = QLabel("Seleccionar Base de Datos")
        title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        self.connection_info_label = QLabel("")
        self.connection_info_label.setObjectName("successLabel")
        self.connection_info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.connection_info_label)

        card = QGroupBox("Bases de Datos Disponibles")
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(18)

        self.database_combo = QComboBox()
        card_layout.addWidget(self.database_combo)

        self.db_status_label = QLabel("Cargando bases de datos...")
        self.db_status_label.setObjectName("infoLabel")
        card_layout.addWidget(self.db_status_label)

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(12)

        self.select_db_btn = QPushButton("Seleccionar Base de Datos")
        self.select_db_btn.setEnabled(False)
        self.select_db_btn.clicked.connect(self.select_database)
        buttons_layout.addWidget(self.select_db_btn)

        back_btn = QPushButton("← Cambiar Conexión")
        back_btn.clicked.connect(self.back_to_connection)
        buttons_layout.addWidget(back_btn)

        card_layout.addLayout(buttons_layout)

        self.final_status_label = QLabel("")
        self.final_status_label.setObjectName("infoLabel")
        card_layout.addWidget(self.final_status_label)

        layout.addWidget(card)
        layout.addStretch()
        self.stacked.addWidget(self.database_page)

    def setup_export_page(self):
        self.export_page = QWidget()
        page_layout = QVBoxLayout(self.export_page)
        page_layout.setContentsMargins(24, 24, 24, 24)
        page_layout.setSpacing(24)

        # Logo centrado
        logo_label = self._create_logo_label(80)
        if logo_label:
            logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            page_layout.addWidget(logo_label, alignment=Qt.AlignmentFlag.AlignCenter)

        title = QLabel("Exportar Datos a TXT")
        title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        page_layout.addWidget(title)

        self.export_info_label = QLabel("")
        self.export_info_label.setObjectName("successLabel")
        self.export_info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        page_layout.addWidget(self.export_info_label)

        # Sección MAESTROS
        maestros_group = QGroupBox("📋 MAESTROS")
        maestros_layout = QVBoxLayout(maestros_group)
        maestros_layout.setSpacing(16)

        self.maestros_grid = QGridLayout()
        self.maestros_grid.setHorizontalSpacing(20)
        self.maestros_grid.setVerticalSpacing(10)

        self.maestros_checkboxes = {}
        # SOLO los nombres de las tablas, sin descripciones
        maestros_tables = [
            "Artículos",
            "Categoría",
            "Control Sanitario",
            "Marcas",
            "Usos",
            "Proveedores",
            "Principios Activos",
            "Bancos",
            "Forma de Pago",
        ]

        for index, table in enumerate(maestros_tables):
            checkbox = QCheckBox(table)
            checkbox.setChecked(True)
            
            row = index // 3  # 3 columnas
            col = index % 3
            self.maestros_grid.addWidget(checkbox, row, col, alignment=Qt.AlignmentFlag.AlignLeft)
            self.maestros_checkboxes[table] = checkbox

        maestros_layout.addLayout(self.maestros_grid)

        maestros_buttons_layout = QHBoxLayout()
        maestros_buttons_layout.setSpacing(12)
        maestros_select_all_btn = QPushButton("✅ Seleccionar Todo")
        maestros_select_all_btn.clicked.connect(lambda: self.select_all_tables("maestros"))
        maestros_deselect_all_btn = QPushButton("❌ Deseleccionar Todo")
        maestros_deselect_all_btn.clicked.connect(lambda: self.deselect_all_tables("maestros"))
        maestros_buttons_layout.addWidget(maestros_select_all_btn)
        maestros_buttons_layout.addWidget(maestros_deselect_all_btn)
        maestros_buttons_layout.addStretch()
        maestros_layout.addLayout(maestros_buttons_layout)

        page_layout.addWidget(maestros_group)

        # Sección RELACIONES
        relaciones_group = QGroupBox("🔗 RELACIONES")
        relaciones_layout = QVBoxLayout(relaciones_group)
        relaciones_layout.setSpacing(16)

        self.relaciones_grid = QGridLayout()
        self.relaciones_grid.setHorizontalSpacing(20)
        self.relaciones_grid.setVerticalSpacing(10)

        self.relaciones_checkboxes = {}
        # SOLO los nombres de las relaciones, sin descripciones
        relaciones_tables = [
            "Artículos - Categorías",
            "Artículos - Códigos de Barras",
            "Artículos - Componentes",
            "Artículos - Control Sanitario",
            "Artículos - Marcas",
            "Artículos - Principio Activo",
            "Artículos - Unidades de Medida",
            "Artículos - Usos",
            "Artículos - Impuesto",
            "Artículos - Atributos (Medicina)",
            "Artículos - Atributos (Genérico)",
        ]

        for index, table in enumerate(relaciones_tables):
            checkbox = QCheckBox(table)
            checkbox.setChecked(True)
            
            row = index // 3  # 3 columnas
            col = index % 3
            self.relaciones_grid.addWidget(checkbox, row, col, alignment=Qt.AlignmentFlag.AlignLeft)
            self.relaciones_checkboxes[table] = checkbox

        relaciones_layout.addLayout(self.relaciones_grid)

        relaciones_buttons_layout = QHBoxLayout()
        relaciones_buttons_layout.setSpacing(12)
        relaciones_select_all_btn = QPushButton("✅ Seleccionar Todo")
        relaciones_select_all_btn.clicked.connect(lambda: self.select_all_tables("relaciones"))
        relaciones_deselect_all_btn = QPushButton("❌ Deseleccionar Todo")
        relaciones_deselect_all_btn.clicked.connect(lambda: self.deselect_all_tables("relaciones"))
        relaciones_buttons_layout.addWidget(relaciones_select_all_btn)
        relaciones_buttons_layout.addWidget(relaciones_deselect_all_btn)
        relaciones_buttons_layout.addStretch()
        relaciones_layout.addLayout(relaciones_buttons_layout)

        page_layout.addWidget(relaciones_group)

        destination_group = QGroupBox("📁 Carpeta de Destino")
        destination_layout = QHBoxLayout(destination_group)
        destination_layout.setSpacing(12)

        self.folder_path = os.path.join(os.path.expanduser("~"), "Exportacion_SQL")
        self.folder_entry = QLineEdit(self.folder_path)
        self.folder_entry.setReadOnly(True)
        destination_layout.addWidget(self.folder_entry)

        browse_btn = QPushButton("Examinar…")
        browse_btn.clicked.connect(self.browse_folder)
        destination_layout.addWidget(browse_btn)

        page_layout.addWidget(destination_group)

        progress_group = QGroupBox("🚀 Progreso de Exportación")
        progress_layout = QVBoxLayout(progress_group)
        progress_layout.setSpacing(16)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        progress_layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("Listo para exportar")
        self.progress_label.setObjectName("infoLabel")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        progress_layout.addWidget(self.progress_label)

        self.export_status_label = QLabel("")
        self.export_status_label.setObjectName("successLabel")
        self.export_status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        progress_layout.addWidget(self.export_status_label)

        query_label = QLabel("Consultas ejecutadas:")
        query_label.setObjectName("sectionLabel")
        progress_layout.addWidget(query_label)

        self.query_text = QTextEdit()
        self.query_text.setReadOnly(True)
        progress_layout.addWidget(self.query_text)

        self.export_btn = QPushButton("🎯 Iniciar Exportación")
        self.export_btn.clicked.connect(self.execute_and_export)
        progress_layout.addWidget(self.export_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        page_layout.addWidget(progress_group)

        back_btn = QPushButton("← Cambiar Base de Datos")
        back_btn.clicked.connect(self.back_to_database_selection)
        page_layout.addWidget(back_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        page_layout.addStretch()
        self.stacked.addWidget(self.export_page)

    # ----------- Métodos auxiliares ----------- #
    def _create_logo_label(self, height=60):
        if not self.logo_pixmap:
            return None
        label = QLabel()
        label.setPixmap(
            self.logo_pixmap.scaledToHeight(height, Qt.TransformationMode.SmoothTransformation)
        )
        label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        return label

    def apply_window_icon(self):
        if self.logo_pixmap:
            self.setWindowIcon(QIcon(self.logo_pixmap))

    def apply_stylesheet(self):
        self.setStyleSheet(
            """
            QWidget {
                background: #181818;
                color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 10.5pt;
            }
            QScrollArea {
                border: none;
            }
            QGroupBox {
                background: #1f1f1f;
                border: 1px solid #2d2d2d;
                border-radius: 10px;
                padding: 18px;
            }
            QPushButton {
                background: #3a86ff;
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 9px 18px;
                font-weight: 600;
            }
            QPushButton:disabled {
                background: #3c3c3c;
                color: #9e9e9e;
            }
            QPushButton:hover:!disabled {
                background: #2f6ddb;
            }
            QPushButton:pressed:!disabled {
                background: #2556af;
            }
            QLineEdit, QComboBox, QTextEdit {
                background: #222222;
                border: 1px solid #333333;
                border-radius: 6px;
                padding: 7px;
                selection-background-color: #3a86ff;
                selection-color: #ffffff;
            }
            QComboBox QAbstractItemView {
                background: #222222;
                border: 1px solid #333333;
                selection-background-color: #3a86ff;
                selection-color: #ffffff;
            }
            QTextEdit {
                min-height: 120px;
            }
            QProgressBar {
                background: #222222;
                border: 1px solid #333333;
                border-radius: 6px;
                text-align: center;
            }
            QProgressBar::chunk {
                background: #3a86ff;
                border-radius: 6px;
            }
            QCheckBox {
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 1px solid #3a3a3a;
                background: #222222;
            }
            QCheckBox::indicator:checked {
                background: #3a86ff;
                border: 1px solid #3a86ff;
            }
            #titleLabel {
                font-size: 20pt;
                font-weight: 700;
            }
            #subtitleLabel {
                font-size: 11pt;
                font-weight: 600;
                color: #b0bec5;
            }
            #sectionLabel {
                font-size: 11pt;
                font-weight: 600;
                color: #cfd8dc;
            }
            #infoLabel {
                color: #64b5f6;
                font-size: 10pt;
            }
            #successLabel {
                color: #81c784;
                font-size: 10pt;
            }
            #hintLabel {
                color: #a0a0a0;
                font-size: 9.5pt;
            }
            QScrollBar:vertical {
                background: #1c1c1c;
                width: 12px;
                margin: 6px 0 6px 0;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #2f2f2f;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical:hover {
                background: #3a3a3a;
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0;
            }
            """
        )

    def toggle_credentials(self):
        enabled = not self.windows_check.isChecked()
        self.user_entry.setEnabled(enabled)
        self.pass_entry.setEnabled(enabled)

    def _get_local_sql_instances(self):
        instances = []
        hostname = socket.gethostname()
        registry_paths = [
            r"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL",
            r"SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL",
        ]
        for path in registry_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                    index = 0
                    while True:
                        try:
                            name, _, _ = winreg.EnumValue(key, index)
                            if name.upper() == "MSSQLSERVER":
                                instances.append(hostname)
                            else:
                                instances.append(f"{hostname}\\{name}")
                            index += 1
                        except OSError:
                            break
            except FileNotFoundError:
                continue
            except Exception as exc:  # noqa: F841
                print(f"Info: No se pudieron leer instancias locales: {exc}")
        # Remover duplicados manteniendo orden
        seen = set()
        ordered = []
        for instance in instances:
            if instance not in seen:
                ordered.append(instance)
                seen.add(instance)
        return ordered

    def _is_instance_reachable(self, instance_name):
        conn_str = (
            "DRIVER={ODBC Driver 17 for SQL Server};"
            f"SERVER={instance_name};"
            "Trusted_Connection=yes;"
            "Connection Timeout=2;"
        )
        try:
            conn = pyodbc.connect(conn_str)
            conn.close()
            return True
        except pyodbc.Error as exc:
            error_msg = str(exc)
            if "Login failed" in error_msg or "Error de autenticación" in error_msg:
                return True
            if "network-related" in error_msg.lower() or "server does not exist" in error_msg.lower():
                return False
            if "timeout expired" in error_msg.lower():
                return False
            return False
        except Exception:
            return False

    def _filter_valid_instances(self, candidates):
        valid_instances = []
        for instance_name in candidates:
            if self._is_instance_reachable(instance_name):
                valid_instances.append(instance_name)
        return valid_instances

    def search_instances(self):
        self.search_btn.setEnabled(False)
        self.status_label.setText("Buscando instancias...")
        QApplication.processEvents()
        
        try:
            instances = []
            hostname = socket.gethostname()
            local_instances = self._get_local_sql_instances()

            if local_instances:
                instances.extend(local_instances)
            else:
                common = [
                    f"{hostname}",
                    f"{hostname}\\SQLEXPRESS",
                    "localhost",
                    "localhost\\SQLEXPRESS",
                    ".",
                    ".\\SQLEXPRESS",
                    "(local)",
                    "(local)\\SQLEXPRESS",
                ]
                instances.extend(common)

            try:
                sources = pyodbc.dataSources()
                for key in sources.keys():
                    if "SQL Server" in sources[key] or "SQL Server" in key:
                        instances.append(key)
            except Exception as exc:  # noqa: F841 - solo informativo
                print(f"Info: No se pudieron obtener fuentes de datos: {exc}")

            instances = sorted(set(instances))

            if not local_instances:
                valid_instances = self._filter_valid_instances(instances)
                if valid_instances:
                    instances = valid_instances

            self.instance_combo.clear()
            self.instance_combo.addItems(instances)
            if instances:
                self.instance_combo.setCurrentIndex(0)
                self.connect_btn.setEnabled(True)
                self.status_label.setText(f"✅ Se encontraron {len(instances)} instancia(s)")
            else:
                self.status_label.setText("⚠️ No se encontraron instancias. Ingresa manualmente.")
                self.connect_btn.setEnabled(True)
        except Exception as exc:
            self.status_label.setText("❌ Error en la búsqueda")
            QMessageBox.critical(self, "Error", f"Error al buscar instancias:\n{exc}")
        finally:
            self.search_btn.setEnabled(True)

    def connect_to_server(self):
        instance = self.instance_combo.currentText().strip()
        if not instance:
            QMessageBox.warning(self, "Advertencia", "Por favor selecciona una instancia")
            return
        
        try:
            if self.windows_check.isChecked():
                conn_str = (
                    "DRIVER={ODBC Driver 17 for SQL Server};"
                    f"SERVER={instance};"
                    "Trusted_Connection=yes;"
                    "Connection Timeout=10;"
                )
            else:
                username = self.user_entry.text().strip()
                password = self.pass_entry.text()
                if not username or not password:
                    QMessageBox.warning(self, "Advertencia", "Por favor ingresa usuario y contraseña")
                    return
                conn_str = (
                    "DRIVER={ODBC Driver 17 for SQL Server};"
                    f"SERVER={instance};"
                    f"UID={username};"
                    f"PWD={password};"
                    "Connection Timeout=10;"
                )

            self.status_label.setText("Conectando…")
            self.connect_btn.setEnabled(False)
            QApplication.processEvents()

            self.connection = pyodbc.connect(conn_str)
            self.connected_instance = instance
            self.connection_info_label.setText(f"Conectado a: {instance}")
            self.stacked.setCurrentWidget(self.database_page)

            QTimer.singleShot(100, self.load_databases)
        except pyodbc.Error as exc:
            error_msg = str(exc)
            if "Login failed" in error_msg:
                error_msg = "Error de autenticación: Usuario o contraseña incorrectos"
            elif "timeout" in error_msg.lower():
                error_msg = "Timeout: No se pudo conectar al servidor"
            QMessageBox.critical(self, "Error de Conexión", f"No se pudo conectar:\n\n{error_msg}")
            self.connect_btn.setEnabled(True)
            self.status_label.setText("❌ Error de conexión")
        except Exception as exc:
            QMessageBox.critical(self, "Error", f"Error inesperado:\n{exc}")
            self.connect_btn.setEnabled(True)
            self.status_label.setText("❌ Error inesperado")

    def load_databases(self):
        self.database_combo.clear()
        self.db_status_label.setText("Cargando bases de datos...")
        self.select_db_btn.setEnabled(False)
        QApplication.processEvents()

        try:
            cursor = self.connection.cursor()
            queries = [
                "SELECT name FROM sys.databases WHERE state_desc='ONLINE' AND database_id>4 ORDER BY name",
                "SELECT name FROM sys.databases WHERE state_desc='ONLINE' ORDER BY name",
                "SELECT name FROM sys.databases ORDER BY name",
            ]
            databases = []
            for query in queries:
                try:
                    cursor.execute(query)
                    rows = cursor.fetchall()
                    if rows:
                        databases = [str(row[0]) for row in rows]
                        break
                except Exception:
                    continue
            
            cursor.close()
            if databases:
                self.database_combo.addItems(databases)
                self.database_combo.setCurrentIndex(0)
                self.select_db_btn.setEnabled(True)
                self.db_status_label.setText(f"✅ Se encontraron {len(databases)} base(s) de datos")
            else:
                self.db_status_label.setText("⚠️ No se encontraron bases de datos")
        except Exception as exc:
            self.db_status_label.setText("❌ Error al cargar bases de datos")
            QMessageBox.critical(self, "Error", f"No se pudieron cargar las bases de datos:\n{exc}")

    def select_database(self):
        database = self.database_combo.currentText()
        if not database:
            QMessageBox.warning(self, "Advertencia", "Por favor selecciona una base de datos")
            return
        
        try:
            self.final_status_label.setText("Conectando a la base de datos…")
            self.select_db_btn.setEnabled(False)
            QApplication.processEvents()
            
            cursor = self.connection.cursor()
            cursor.execute(f"USE [{database}]")
            cursor.close()
            
            self.selected_database = database
            self.final_status_label.setText(f"✅ Conectado exitosamente a: {database}")

            def switch_to_export():
                self.export_info_label.setText(
                    f"Conectado a: {self.connected_instance} → {self.selected_database}"
                )
                if not os.path.exists(self.folder_path):
                    os.makedirs(self.folder_path, exist_ok=True)
                self.folder_entry.setText(self.folder_path)
                self.query_text.clear()
                self.progress_bar.setValue(0)
                self.progress_label.setText("Listo para exportar")
                self.export_status_label.clear()
                self.stacked.setCurrentWidget(self.export_page)

            QTimer.singleShot(600, switch_to_export)
        except Exception as exc:
            self.final_status_label.setText("❌ Error al seleccionar base de datos")
            self.select_db_btn.setEnabled(True)
            QMessageBox.critical(self, "Error", f"No se pudo seleccionar la base de datos:\n{exc}")

    # ----------- Exportación ----------- #
    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Seleccionar Carpeta", self.folder_path, QFileDialog.Option.ShowDirsOnly
        )
        if folder:
            self.folder_path = folder
            self.folder_entry.setText(folder)
            if not os.path.exists(folder):
                os.makedirs(folder, exist_ok=True)

    def select_all_tables(self, section):
        if section == "maestros":
            for checkbox in self.maestros_checkboxes.values():
                checkbox.setChecked(True)
        elif section == "relaciones":
            for checkbox in self.relaciones_checkboxes.values():
                checkbox.setChecked(True)

    def deselect_all_tables(self, section):
        if section == "maestros":
            for checkbox in self.maestros_checkboxes.values():
                checkbox.setChecked(False)
        elif section == "relaciones":
            for checkbox in self.relaciones_checkboxes.values():
                checkbox.setChecked(False)

    def log_query(self, table, query):
        if self.query_text.toPlainText():
            self.query_text.append("\n")
        self.query_text.append(f"[{table}] {query.strip()}")
        cursor = self.query_text.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.query_text.setTextCursor(cursor)

    def _normalize_text(self, value):
        """Elimina tildes, caracteres especiales y convierte ñ a n"""
        if value is None:
            return ""
        
        value_str = str(value)
        
        # Si está vacío, retornar vacío
        if not value_str.strip():
            return value_str
        
        # Si es un número o contiene solo dígitos, puntos, comas y signos, no normalizar
        cleaned = value_str.replace('.', '').replace(',', '').replace('-', '').replace('+', '').strip()
        if cleaned.isdigit() or cleaned == '':
            return value_str
        
        # Verificar si contiene letras (texto), si no contiene letras no normalizar
        has_letters = any(char.isalpha() for char in value_str)
        if not has_letters:
            return value_str
        
        # Primero convertir ñ y Ñ a n y N
        value_str = value_str.replace('ñ', 'n').replace('Ñ', 'N')
        
        # Eliminar caracteres especiales como acentos graves y apostrofos
        value_str = value_str.replace('´', '').replace('`', '').replace("'", "").replace("‘", "").replace("’", "")
        
        # Normalizar y eliminar diacríticos (tildes, acentos, etc.) solo para texto
        nfd = unicodedata.normalize('NFD', value_str)
        # Filtrar solo los caracteres que no son marcas diacríticas (combining characters)
        normalized = ''.join(char for char in nfd if unicodedata.category(char) != 'Mn')
        
        return normalized

    def _clean_special_characters(self, text):
        """Elimina caracteres especiales problemáticos manteniendo la legibilidad"""
        if not text:
            return text
        
        # Caracteres problemáticos comunes en archivos de texto
        problematic_chars = {
            '\x00': '',  # Null character
            '\x01': '',  # Start of Header
            '\x02': '',  # Start of Text
            '\x03': '',  # End of Text
            '\x04': '',  # End of Transmission
            '\x05': '',  # Enquiry
            '\x06': '',  # Acknowledge
            '\x07': '',  # Bell
            '\x08': '',  # Backspace
            '\x0b': '',  # Vertical Tab
            '\x0c': '',  # Form Feed
            '\x0e': '',  # Shift Out
            '\x0f': '',  # Shift In
            '\x10': '',  # Data Link Escape
            '\x11': '',  # Device Control 1
            '\x12': '',  # Device Control 2
            '\x13': '',  # Device Control 3
            '\x14': '',  # Device Control 4
            '\x15': '',  # Negative Acknowledge
            '\x16': '',  # Synchronous Idle
            '\x17': '',  # End of Transmission Block
            '\x18': '',  # Cancel
            '\x19': '',  # End of Medium
            '\x1a': '',  # Substitute
            '\x1b': '',  # Escape
            '\x1c': '',  # File Separator
            '\x1d': '',  # Group Separator
            '\x1e': '',  # Record Separator
            '\x1f': '',  # Unit Separator
            '\x7f': '',  # Delete
            '\ufffe': '', # BOM
            '\uffff': '', # BOM
            '´': '',     # Acento agudo
            '`': '',     # Acento grave
            '‘': '',     # Apostrofo curvo izquierdo
            '’': '',     # Apostrofo curvo derecho
            "'": "",     # Apostrofo simple
        }
        
        cleaned_text = text
        for char, replacement in problematic_chars.items():
            cleaned_text = cleaned_text.replace(char, replacement)
        
        return cleaned_text

    def execute_and_export(self):
        # Combinar selecciones de ambas secciones
        selected_maestros = [name for name, cb in self.maestros_checkboxes.items() if cb.isChecked()]
        selected_relaciones = [name for name, cb in self.relaciones_checkboxes.items() if cb.isChecked()]
        
        selected_tables = selected_maestros + selected_relaciones
        if not selected_tables:
            QMessageBox.warning(self, "Advertencia", "Selecciona al menos una tabla o relación")
            return

        folder = self.folder_entry.text()
        if not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)

        self.export_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_label.setText("Iniciando exportación…")
        self.export_status_label.clear()
        self.query_text.clear()
        QApplication.processEvents()

        total_tables = len(selected_tables)
        exported_files = []
        total_records = 0

        for index, table in enumerate(selected_tables, start=1):
            progress = int(((index - 1) / total_tables) * 100)
            self.progress_bar.setValue(progress)
            self.progress_label.setText(f"Exportando {table}… ({index}/{total_tables})")
            QApplication.processEvents()

            query = self.get_query_for_table(table)
            if not query:
                continue

            self.log_query(table, query)

            try:
                cursor = self.connection.cursor()
                cursor.execute(query)
                rows = cursor.fetchall()

                # Generar nombre de archivo limpiando caracteres especiales
                filename = table.replace(' ', '_').replace('-', '_').replace('(', '').replace(')', '').lower()
                filename = f"{filename}.txt"
                filepath = os.path.join(folder, filename)

                with open(filepath, "w", encoding="utf-8") as f:
                    for row in rows:
                        # Normalizar solo los valores de texto (quitar tildes y caracteres especiales)
                        normalized_row = []
                        for value in row:
                            normalized_value = self._normalize_text(value)
                            # Limpiar caracteres especiales problemáticos
                            cleaned_value = self._clean_special_characters(normalized_value)
                            normalized_row.append(cleaned_value)
                        
                        line = ";".join(str(value) for value in normalized_row)
                        f.write(line + "\n")

                cursor.close()
                records_count = len(rows)
                total_records += records_count
                exported_files.append(f"• {filename} ({records_count} registros)")
                self.export_status_label.setText(f"✓ {table}: {records_count} registros")
            except Exception as exc:
                self.export_status_label.setText(f"✗ {table}: Error ({exc})")
                continue

        self.progress_bar.setValue(100)
        self.progress_label.setText("Exportación completada")
        self.export_btn.setEnabled(True)

        summary = "\n".join(exported_files)
        message = (
            "✅ Exportación completada!\n\n"
            f"Tablas: {len(exported_files)}/{total_tables}\n"
            f"Registros: {total_records}\n"
            f"Ubicación: {folder}\n\n"
            f"Archivos:\n{summary}"
        )
        reply = QMessageBox.question(self, "Éxito", f"{message}\n\n¿Abrir carpeta?")
        if reply == QMessageBox.StandardButton.Yes:
            try:
                os.startfile(folder)
            except Exception:
                # Para sistemas no Windows
                import subprocess
                import platform
                system = platform.system()
                if system == "Darwin":  # macOS
                    subprocess.run(["open", folder])
                else:  # Linux
                    subprocess.run(["xdg-open", folder])

    def get_query_for_table(self, table_name):
        # Consultas MAESTROS (las originales)
        maestros_queries = {
            "Artículos": "SELECT CodigoArticulo, REPLACE(Descripcion,';',' '), "
            "REPLACE(DescripcionLarga,';',' '), CPE, PermisoSanitario, "
            "CAST(ValorMaximo as int), CAST(ValorMinimo as int), "
            "CASE WHEN DerechoOferta IS NULL THEN 'NULL' ELSE CAST(DerechoOferta as int) END, "
            "CAST(GenerarEtiquetaCompras as int) "
            "FROM InvArticulo JOIN InvMinMax on InvMinMax.InvArticuloId = InvArticulo.Id",
            "Categoría": "SELECT CodigoCategoria, REPLACE(Descripcion,';',' ') FROM InvCategoria",
            "Control Sanitario": "SELECT CodigoControl, REPLACE(Descripcion,';',' ') FROM InvControlSanitario",
            "Marcas": "SELECT Codigo, REPLACE(Nombre,';',' '), "
            "CASE WHEN Nota IS NULL THEN 'NULL' ELSE REPLACE(Nota,';',' ') END FROM InvMarca",
            "Usos": "SELECT Codigo, REPLACE(Descripcion,';',' ') FROM InvUso",
            "Proveedores": """WITH ProveedoresLimpios AS (
    SELECT 
        CodigoProveedor, 
        REPLACE(Nombre,';',' ') AS NombreLimpio,
        CASE WHEN Apellido IS NULL THEN 'NULL' ELSE REPLACE(Apellido,';',' ') END AS ApellidoLimpio,
        REPLACE(IdentificacionFiscal,';',' ') AS IdentificacionFiscalLimpia,
        Telefono, 
        Representante, 
        Contactos,
        CASE WHEN NombreEnCheque IS NULL THEN 'NULL' ELSE REPLACE(NombreEnCheque,';',' ') END AS NombreEnChequeLimpio,
        TipoContribuyente, 
        TipoPersona, 
        tipoProveedor,
        ROW_NUMBER() OVER (
            PARTITION BY 
                REPLACE(IdentificacionFiscal,';',' '),
                UPPER(REPLACE(REPLACE(Nombre,';',' '), ' ', ''))
            ORDER BY CodigoProveedor ASC
        ) AS FilaNum
    FROM ComProveedor AS cp 
    JOIN GenPersona AS g ON g.Id = cp.GenPersonaId 
)
SELECT 
    CodigoProveedor, 
    NombreLimpio,
    ApellidoLimpio,
    IdentificacionFiscalLimpia,
    Telefono, 
    Representante, 
    Contactos,
    NombreEnChequeLimpio,
    TipoContribuyente, 
    TipoPersona, 
    tipoProveedor
FROM ProveedoresLimpios
WHERE FilaNum = 1
ORDER BY CodigoProveedor ASC""",
            "Principios Activos": "SELECT Codigo, REPLACE(Nombre,';',' '), CAST(SustanciaControlada as int) "
            "FROM InvComponente",
            "Bancos": "SELECT Codigo, REPLACE(Descripcion,';',' '), REPLACE(Nombre,';',' '), "
            "tipoConfiguracion, estado FROM BanBanco",
            "Forma de Pago": "SELECT Codigo, Nombre, Descripcion, estado, tipoFormaPago, "
            "AplicaRetencion, PorcentajeRetencion, AplicaComisionFormaPago, PorcentajeComisionFormaPago, "
            "CodigoImpresoraFiscal, AplicaIva "
            "FROM BanFormaPago WHERE Codigo IN ('FP01','FP02','FP03','FP04','FP05','FP06','FP07')",
        }

        # Consultas RELACIONES (las nuevas)
        relaciones_queries = {
            "Artículos - Categorías": "SELECT CodigoArticulo, CodigoCategoria FROM InvArticulo JOIN InvCategoria on InvCategoria.Id = InvArticulo.InvCategoriaId",
            "Artículos - Códigos de Barras": """WITH CodigosBarrasLimpios AS (
    SELECT 
        CodigoArticulo, 
        CodigoBarra, 
        CAST(EsPrincipal as int) AS EsPrincipal,
        ROW_NUMBER() OVER (
            PARTITION BY CodigoBarra 
            ORDER BY EsPrincipal DESC, CodigoArticulo ASC
        ) AS FilaNum
    FROM InvArticulo 
    JOIN InvCodigoBarra on InvCodigoBarra.InvArticuloId = InvArticulo.Id
    WHERE CodigoBarra IS NOT NULL AND CodigoBarra != ''
)
SELECT CodigoArticulo, CodigoBarra, EsPrincipal
FROM CodigosBarrasLimpios
WHERE FilaNum = 1""",
            "Artículos - Componentes": "select InvArticulo.CodigoArticulo, InvComponente.Codigo from InvArticuloComponente join InvComponente on InvComponente.Id = InvArticuloComponente.InvComponenteId join InvArticulo on InvArticulo.Id = InvArticuloComponente.InvArticuloId",
            "Artículos - Control Sanitario": "SELECT CodigoArticulo, InvControlSanitario.CodigoControl FROM InvArticulo JOIN InvControlSanitario on InvControlSanitario.Id = InvArticulo.InvControlSanitarioId",
            "Artículos - Marcas": "SELECT CodigoArticulo, Codigo FROM InvArticulo JOIN InvMarca on InvMarca.Id = InvArticulo.InvMarcaId",
            "Artículos - Principio Activo": "SELECT CodigoArticulo, InvComponente.Codigo FROM InvArticuloComponente JOIN InvArticulo on InvArticulo.Id = InvArticuloComponente.InvArticuloId JOIN InvComponente on InvComponente.Id = InvArticuloComponente.InvComponenteId",
            "Artículos - Unidades de Medida": "SELECT CodigoArticulo, 'UN', 'UFF', CAST(FactorConversion as int), 'UN', 1 "
            "FROM InvArticuloUnidad UN JOIN InvArticulo ART ON UN.InvArticuloId = ART.Id "
            "WHERE FactorConversion > 1",
            "Artículos - Usos": "SELECT CodigoArticulo, InvUso.Codigo FROM InvArticuloUso JOIN InvArticulo on InvArticulo.Id = InvArticuloUso.InvArticuloId JOIN InvUso on InvUso.Id = InvArticuloUso.InvUsoId",
            "Artículos - Impuesto": "SELECT ex.CodigoArticulo, CAST(MAX(ex.tarifaI) AS NUMERIC(10,2)) as TarifaCompra, CAST(MAX(ex.tarifaV) AS NUMERIC(10,2)) as TarifaVenta FROM ( SELECT a.CodigoArticulo, CASE WHEN MAX(fcv.TarifaImpuesto) IS NULL THEN 0 ELSE CAST(MAX(fcv.TarifaImpuesto) AS DECIMAL(10,2)) END as tarifaI, CASE WHEN MAX(fcv2.TarifaImpuesto) IS NULL THEN 0 ELSE CAST(MAX(fcv2.TarifaImpuesto) AS DECIMAL(10,2)) END as tarifaV FROM InvArticulo a LEFT JOIN FinConceptoVigencia fcv ON fcv.FinConceptoImptoId = a.FinConceptoImptoIdCompra LEFT JOIN FinConceptoVigencia fcv2 ON fcv2.FinConceptoImptoId = a.FinConceptoImptoIdVenta GROUP BY a.CodigoArticulo ) as ex GROUP BY ex.CodigoArticulo",
            "Artículos - Atributos (Medicina)": "SELECT InvArticulo.CodigoArticulo FROM InvArticulo JOIN InvArticuloAtributo ON InvArticuloAtributo.InvArticuloId = InvArticulo.Id JOIN InvAtributo ON InvAtributo.Id = InvArticuloAtributo.InvAtributoId WHERE InvAtributo.Descripcion = 'Medicina' ORDER BY InvArticulo.CodigoArticulo ASC",
            "Artículos - Atributos (Genérico)": "SELECT InvArticulo.CodigoArticulo FROM InvArticulo JOIN InvArticuloAtributo ON InvArticuloAtributo.InvArticuloId = InvArticulo.Id JOIN InvAtributo ON InvAtributo.Id = InvArticuloAtributo.InvAtributoId WHERE InvAtributo.Descripcion = 'Genérico' ORDER BY InvArticulo.CodigoArticulo ASC",
        }

        # Buscar en ambas secciones
        if table_name in maestros_queries:
            return maestros_queries[table_name]
        elif table_name in relaciones_queries:
            return relaciones_queries[table_name]
        else:
            return None

    # ----------- Navegación ----------- #
    def back_to_database_selection(self):
        self.query_text.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText("Listo para exportar")
        self.export_status_label.clear()
        self.stacked.setCurrentWidget(self.database_page)

    def back_to_connection(self):
        if self.connection:
            self.connection.close()
            self.connection = None
        self.select_db_btn.setEnabled(False)
        self.instance_combo.clear()
        self.status_label.setText("Listo para conectar")
        self.final_status_label.clear()
        self.db_status_label.setText("Cargando bases de datos...")
        self.connected_instance = ""
        self.selected_database = ""
        self.stacked.setCurrentWidget(self.connection_page)

    def closeEvent(self, event):  # noqa: N802
        if self.connection:
            try:
                self.connection.close()
            except Exception:
                pass
        event.accept()


def main():
    app = QApplication([])
    window = SQLServerConnector()
    window.showMaximized()
    app.exec()


if __name__ == "__main__":
    main()