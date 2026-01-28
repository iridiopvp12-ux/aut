import flet as ft
import threading
from pathlib import Path
from src.logic.fiscal_logic import executar_analise_completa
import logging

class SpedView(ft.Container):
    def __init__(self, page: ft.Page):
        super().__init__()
        self.page = page
        self.expand = True
        self.padding = 20

        # Variables to store paths
        self.sped_path_val = None
        self.xml_folder_val = None
        self.rules_path_val = None
        self.detailed_rules_path_val = None
        self.template_path_val = None

        # --- UI Components ---

        # 1. SPED File Selection
        self.sped_path_text = ft.Text(value="Nenhum arquivo selecionado", italic=True)
        self.pick_sped_dialog = ft.FilePicker(on_result=self.pick_sped_result)
        self.page.overlay.append(self.pick_sped_dialog)

        # 2. XML Folder Selection
        self.xml_path_text = ft.Text(value="Nenhuma pasta selecionada", italic=True)
        self.pick_xml_dialog = ft.FilePicker(on_result=self.pick_xml_result)
        self.page.overlay.append(self.pick_xml_dialog)

        # 3. Rules File Selection
        self.rules_path_text = ft.Text(value="Nenhum arquivo selecionado", italic=True)
        self.pick_rules_dialog = ft.FilePicker(on_result=self.pick_rules_result)
        self.page.overlay.append(self.pick_rules_dialog)

        # 4. Sector Selection
        self.sector_dropdown = ft.Dropdown(
            label="Setor / Atividade",
            options=[
                ft.dropdown.Option("Comercio"),
                ft.dropdown.Option("Moveleiro"),
                ft.dropdown.Option("E-commerce"),
            ],
            value="Comercio",
            width=200
        )

        # 5. Detailed Rules (Optional)
        self.detailed_rules_checkbox = ft.Checkbox(label="Usar Regras Detalhadas (NCM)?", on_change=self.toggle_detailed_rules)
        self.detailed_rules_container = ft.Column(visible=False)
        self.detailed_rules_text = ft.Text(value="Nenhum arquivo selecionado", italic=True)
        self.pick_detailed_rules_dialog = ft.FilePicker(on_result=self.pick_detailed_rules_result)
        self.page.overlay.append(self.pick_detailed_rules_dialog)
        self.detailed_rules_container.controls.append(
            ft.Row([
                ft.ElevatedButton("Selecionar Regras Detalhadas", icon=ft.Icons.UPLOAD_FILE, on_click=lambda _: self.pick_detailed_rules_dialog.pick_files(allow_multiple=False, allowed_extensions=["xlsx", "xls"])),
                self.detailed_rules_text
            ])
        )

        # 6. Template Apuracao (Optional)
        self.template_checkbox = ft.Checkbox(label="Preencher Template de Apuração?", on_change=self.toggle_template)
        self.template_container = ft.Column(visible=False)
        self.template_text = ft.Text(value="Nenhum arquivo selecionado", italic=True)
        self.pick_template_dialog = ft.FilePicker(on_result=self.pick_template_result)
        self.page.overlay.append(self.pick_template_dialog)
        self.template_container.controls.append(
            ft.Row([
                ft.ElevatedButton("Selecionar Template", icon=ft.Icons.UPLOAD_FILE, on_click=lambda _: self.pick_template_dialog.pick_files(allow_multiple=False, allowed_extensions=["xlsx"])),
                self.template_text
            ])
        )

        # Output Area
        self.status_text = ft.Text("Aguardando início...", size=16, weight="bold")
        self.progress_bar = ft.ProgressBar(width=600, value=0, visible=False)
        self.log_view = ft.ListView(expand=True, spacing=5, auto_scroll=True, height=200)
        self.log_container = ft.Container(
            content=self.log_view,
            border=ft.border.all(1, ft.Colors.GREY_400),
            border_radius=5,
            padding=10,
            bgcolor=ft.Colors.BLACK12,
            expand=True
        )

        # Start Button
        self.start_button = ft.ElevatedButton(
            "INICIAR ANÁLISE COMPLETA",
            icon=ft.Icons.PLAY_ARROW,
            style=ft.ButtonStyle(color=ft.Colors.WHITE, bgcolor=ft.Colors.BLUE),
            on_click=self.start_analysis,
            disabled=True
        )

        # Layout
        self.content = ft.Column(
            [
                ft.Text("Analisador Fiscal", size=30, weight="bold"),
                ft.Divider(),

                ft.Row([
                    ft.Column([
                        ft.Text("Arquivo SPED Fiscal:", weight="bold"),
                        ft.Row([
                            ft.ElevatedButton("Selecionar SPED", icon=ft.Icons.UPLOAD_FILE, on_click=lambda _: self.pick_sped_dialog.pick_files(allow_multiple=False, allowed_extensions=["txt"])),
                            self.sped_path_text
                        ])
                    ]),
                    ft.Column([
                        ft.Text("Pasta XMLs:", weight="bold"),
                        ft.Row([
                            ft.ElevatedButton("Selecionar Pasta", icon=ft.Icons.FOLDER_OPEN, on_click=lambda _: self.pick_xml_dialog.get_directory_path()),
                            self.xml_path_text
                        ])
                    ]),
                ]),
                ft.Row([
                    ft.Column([
                        ft.Text("Arquivo de Regras (Acumuladores):", weight="bold"),
                        ft.Row([
                            ft.ElevatedButton("Selecionar Regras", icon=ft.Icons.UPLOAD_FILE, on_click=lambda _: self.pick_rules_dialog.pick_files(allow_multiple=False, allowed_extensions=["xlsx", "xls", "csv"])),
                            self.rules_path_text
                        ])
                    ]),
                     ft.Column([
                        ft.Text("Setor:", weight="bold"),
                        self.sector_dropdown
                    ]),
                ]),

                ft.Divider(),
                self.detailed_rules_checkbox,
                self.detailed_rules_container,

                self.template_checkbox,
                self.template_container,

                ft.Divider(),
                self.start_button,

                ft.Divider(),
                self.status_text,
                self.progress_bar,
                ft.Text("Logs da Análise:", weight="bold"),
                self.log_container
            ],
            scroll=ft.ScrollMode.AUTO
        )

    # --- File Picker Results ---

    def pick_sped_result(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.sped_path_val = e.files[0].path
            self.sped_path_text.value = e.files[0].name
            self.check_can_start()
            self.update()

    def pick_xml_result(self, e: ft.FilePickerResultEvent):
        if e.path:
            self.xml_folder_val = e.path
            self.xml_path_text.value = e.path
            self.check_can_start()
            self.update()

    def pick_rules_result(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.rules_path_val = e.files[0].path
            self.rules_path_text.value = e.files[0].name
            self.check_can_start()
            self.update()

    def pick_detailed_rules_result(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.detailed_rules_path_val = e.files[0].path
            self.detailed_rules_text.value = e.files[0].name
            self.check_can_start()
            self.update()

    def pick_template_result(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.template_path_val = e.files[0].path
            self.template_text.value = e.files[0].name
            self.update()

    # --- UI Logic ---

    def toggle_detailed_rules(self, e):
        self.detailed_rules_container.visible = self.detailed_rules_checkbox.value
        self.check_can_start()
        self.update()

    def toggle_template(self, e):
        self.template_container.visible = self.template_checkbox.value
        self.update()

    def check_can_start(self):
        can_start = (
            self.sped_path_val is not None and
            self.xml_folder_val is not None and
            self.rules_path_val is not None
        )

        if self.detailed_rules_checkbox.value:
            if self.detailed_rules_path_val is None:
                can_start = False

        self.start_button.disabled = not can_start
        self.update()

    # --- Analysis Logic ---

    def add_log(self, message: str):
        self.log_view.controls.append(ft.Text(message, size=12, color=ft.Colors.WHITE, font_family="Consolas"))
        self.update()

    def update_status(self, message: str):
        self.status_text.value = f"Status: {message}"
        self.add_log(f"[STATUS] {message}")
        self.update()

    def update_progress(self, current, total):
        if total > 0:
            val = current / total
            self.progress_bar.value = val
            self.progress_bar.visible = True
        else:
            self.progress_bar.visible = False
        self.update()

    def on_done(self, report_path, issues_count):
        self.status_text.value = f"Concluído! {issues_count} inconsistências. Relatório: {report_path}"
        self.progress_bar.value = 1.0
        self.start_button.disabled = False

        # Show snackbar
        self.page.snack_bar = ft.SnackBar(ft.Text(f"Análise concluída! Relatório salvo em {report_path}"))
        self.page.snack_bar.open = True
        self.page.update()

    def on_error(self, error_msg):
        self.status_text.value = f"ERRO: {error_msg}"
        self.add_log(f"[ERRO] {error_msg}")
        self.start_button.disabled = False
        self.progress_bar.visible = False
        self.update()

    def start_analysis(self, e):
        self.start_button.disabled = True
        self.log_view.controls.clear()
        self.status_text.value = "Iniciando..."
        self.progress_bar.visible = True
        self.progress_bar.value = None # Indeterminate
        self.update()

        # Configs from logic (assuming we can pass them or they are defaults)
        # In fiscal_logic.py, executing_analise_completa takes:
        # cfop_sem_credito_icms: List[str], cfop_sem_credito_ipi: List[str], tolerancia_valor: float
        # I should probably expose these in settings or hardcode for now based on `config.py` from attachment if it exists.
        # Checking `config.py` in attachment:
        # It was in `src/logic/config.py`? No, I copied it. Let's check `src/logic/config.py`.

        # For now I will use some defaults or try to load from `src.logic.config` if I copied it.
        # I copied all files, so `src/logic/config.py` should be there.

        # Let's try to import ConfigLoader or similar if needed.
        # But to keep it simple, I'll define defaults here matching the typical use case.

        cfop_sem_credito_icms = [
            "1556", "2556", "1407", "2407", "1551", "2551", "1406", "2406",
            "1403", "2403", "1653", "2653"
        ]
        cfop_sem_credito_ipi = [
            "1102", "2102", "1403", "2403", "1556", "2556", "1407", "2407",
            "1551", "2551", "1406", "2406", "1653", "2653"
        ]
        tolerancia_valor = 0.03

        # Paths
        sped = Path(self.sped_path_val)
        xml = Path(self.xml_folder_val)
        rules = Path(self.rules_path_val)

        detailed_rules = Path(self.detailed_rules_path_val) if (self.detailed_rules_checkbox.value and self.detailed_rules_path_val) else None
        template = Path(self.template_path_val) if (self.template_checkbox.value and self.template_path_val) else None

        sector = self.sector_dropdown.value
        username = "admin" # TODO: Get from session

        # Run in thread
        t = threading.Thread(
            target=executar_analise_completa,
            args=(
                sped, xml, rules, username,
                cfop_sem_credito_icms, cfop_sem_credito_ipi, tolerancia_valor
            ),
            kwargs={
                "status_callback": self.update_status,
                "progress_callback": self.update_progress,
                "done_callback": self.on_done,
                "error_callback": self.on_error,
                "caminho_regras_detalhadas": detailed_rules,
                "template_apuracao_path": template,
                "tipo_setor": sector
            },
            daemon=True
        )
        t.start()
