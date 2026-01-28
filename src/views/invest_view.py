import flet as ft
import threading
from pathlib import Path
from src.logic.invest_logic import executar_apuracao_invest
import logging

class InvestView(ft.Container):
    def __init__(self, page: ft.Page):
        super().__init__()
        self.page = page
        self.expand = True
        self.padding = 20

        # Variables to store paths
        self.xml_folder_val = None
        self.sete_path_val = None
        self.ncm_path_val = None

        # --- UI Components ---

        # 1. XML Folder Selection
        self.xml_path_text = ft.Text(value="Nenhuma pasta selecionada", italic=True)
        self.pick_xml_dialog = ft.FilePicker(on_result=self.pick_xml_result)
        self.page.overlay.append(self.pick_xml_dialog)

        # 2. Planilha SETE (Optional)
        self.sete_path_text = ft.Text(value="Nenhum arquivo selecionado (Opcional)", italic=True)
        self.pick_sete_dialog = ft.FilePicker(on_result=self.pick_sete_result)
        self.page.overlay.append(self.pick_sete_dialog)

        # 3. NCM Rules CSV (Optional)
        self.ncm_path_text = ft.Text(value="Nenhum arquivo selecionado (Opcional)", italic=True)
        self.pick_ncm_dialog = ft.FilePicker(on_result=self.pick_ncm_result)
        self.page.overlay.append(self.pick_ncm_dialog)

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
            "INICIAR APURAÇÃO INVEST / CONTRIBUIÇÕES",
            icon=ft.Icons.PLAY_ARROW,
            style=ft.ButtonStyle(color=ft.Colors.WHITE, bgcolor=ft.Colors.GREEN),
            on_click=self.start_analysis,
            disabled=True
        )

        # Layout
        self.content = ft.Column(
            [
                ft.Text("Invest & Contribuições", size=30, weight="bold"),
                ft.Text("Apuração de Incentivos Fiscais e Análise PIS/COFINS", size=16),
                ft.Divider(),

                ft.Row([
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
                        ft.Text("Planilha SETE (Base):", weight="bold"),
                        ft.Row([
                            ft.ElevatedButton("Selecionar Planilha", icon=ft.Icons.UPLOAD_FILE, on_click=lambda _: self.pick_sete_dialog.pick_files(allow_multiple=False, allowed_extensions=["xlsx", "xls"])),
                            self.sete_path_text
                        ])
                    ]),
                    ft.Column([
                        ft.Text("Arquivo Regras NCM:", weight="bold"),
                        ft.Row([
                            ft.ElevatedButton("Selecionar CSV/Excel", icon=ft.Icons.UPLOAD_FILE, on_click=lambda _: self.pick_ncm_dialog.pick_files(allow_multiple=False, allowed_extensions=["csv", "xlsx", "xls"])),
                            self.ncm_path_text
                        ])
                    ]),
                ]),

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

    def pick_xml_result(self, e: ft.FilePickerResultEvent):
        if e.path:
            self.xml_folder_val = e.path
            self.xml_path_text.value = e.path
            self.check_can_start()
            self.update()

    def pick_sete_result(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.sete_path_val = e.files[0].path
            self.sete_path_text.value = e.files[0].name
            self.update()

    def pick_ncm_result(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.ncm_path_val = e.files[0].path
            self.ncm_path_text.value = e.files[0].name
            self.update()

    # --- UI Logic ---

    def check_can_start(self):
        can_start = (self.xml_folder_val is not None)
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

    def on_done(self, result_msg):
        self.status_text.value = f"Concluído!"
        self.progress_bar.value = 1.0
        self.add_log(f"[SUCESSO] {result_msg}")
        self.start_button.disabled = False

        # Show snackbar
        self.page.snack_bar = ft.SnackBar(ft.Text(f"Análise concluída!"))
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
        self.progress_bar.value = None
        self.update()

        xml = Path(self.xml_folder_val)
        sete = self.sete_path_val
        ncm = self.ncm_path_val

        # Run in thread
        t = threading.Thread(
            target=self.run_wrapper,
            args=(xml, sete, ncm),
            daemon=True
        )
        t.start()

    def run_wrapper(self, xml, sete, ncm):
        try:
            executar_apuracao_invest(
                pasta_xml=xml,
                caminho_sete=sete,
                caminho_ncm_csv=ncm,
                status_callback=self.update_status,
                progress_callback=self.update_progress,
                done_callback=self.on_done,
                error_callback=self.on_error
            )
        except Exception as e:
            self.on_error(str(e))
