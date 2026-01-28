import flet as ft
import pandas as pd
from pathlib import Path
from src.logic.sped_parser import extrair_dados_sped
from src.logic.xml_parser import processar_pasta_xml
import threading
from datetime import datetime

class KeysExtractorView(ft.Container):
    def __init__(self, page: ft.Page):
        super().__init__()
        self.page = page
        self.expand = True
        self.padding = 20

        self.input_path_val = None
        self.mode = "SPED" # or "XML"

        # UI Components
        self.mode_dropdown = ft.Dropdown(
            label="Origem dos Dados",
            options=[ft.dropdown.Option("SPED"), ft.dropdown.Option("XML")],
            value="SPED",
            on_change=self.on_mode_change,
            width=200
        )

        self.path_text = ft.Text("Nenhum arquivo selecionado", italic=True)
        self.pick_file_dialog = ft.FilePicker(on_result=self.pick_file_result)
        self.pick_folder_dialog = ft.FilePicker(on_result=self.pick_folder_result)
        self.page.overlay.extend([self.pick_file_dialog, self.pick_folder_dialog])

        self.select_btn = ft.ElevatedButton("Selecionar Arquivo SPED", icon=ft.Icons.UPLOAD_FILE, on_click=self.open_picker)

        self.extract_btn = ft.ElevatedButton("Extrair Chaves", icon=ft.Icons.PLAY_ARROW, on_click=self.start_extraction, disabled=True)

        self.status_text = ft.Text("")
        self.progress_bar = ft.ProgressBar(width=400, visible=False)

        self.content = ft.Column([
            ft.Text("Extrator de Chaves", size=30, weight="bold"),
            ft.Divider(),
            self.mode_dropdown,
            ft.Row([self.select_btn, self.path_text]),
            ft.Divider(),
            self.extract_btn,
            self.progress_bar,
            self.status_text
        ])

    def on_mode_change(self, e):
        self.mode = self.mode_dropdown.value
        self.path_text.value = "Nenhum arquivo selecionado"
        self.input_path_val = None
        self.extract_btn.disabled = True

        if self.mode == "SPED":
            self.select_btn.text = "Selecionar Arquivo SPED"
            self.select_btn.icon = ft.Icons.UPLOAD_FILE
        else:
            self.select_btn.text = "Selecionar Pasta XML"
            self.select_btn.icon = ft.Icons.FOLDER_OPEN
        self.update()

    def open_picker(self, e):
        if self.mode == "SPED":
            self.pick_file_dialog.pick_files(allow_multiple=False, allowed_extensions=["txt"])
        else:
            self.pick_folder_dialog.get_directory_path()

    def pick_file_result(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.input_path_val = e.files[0].path
            self.path_text.value = e.files[0].name
            self.extract_btn.disabled = False
            self.update()

    def pick_folder_result(self, e: ft.FilePickerResultEvent):
        if e.path:
            self.input_path_val = e.path
            self.path_text.value = e.path
            self.extract_btn.disabled = False
            self.update()

    def start_extraction(self, e):
        self.extract_btn.disabled = True
        self.progress_bar.visible = True
        self.status_text.value = "Extraindo..."
        self.update()

        t = threading.Thread(target=self.run_extraction, daemon=True)
        t.start()

    def run_extraction(self):
        try:
            df_keys = pd.DataFrame()

            if self.mode == "SPED":
                # Only interested in the first dataframe (headers) or chaves_difal
                # Actually extrair_dados_sped returns: df_sped, df_sped_itens, df_sped_analitico, df_sped_cte, df_chaves_difal
                # df_sped has CHV_NFE. df_sped_cte has CHV_CTE.
                dfs = extrair_dados_sped(Path(self.input_path_val))
                df_nfe = dfs[0]
                df_cte = dfs[3]

                keys_nfe = df_nfe['CHV_NFE'].unique().tolist() if 'CHV_NFE' in df_nfe.columns else []
                keys_cte = df_cte['CHV_CTE'].unique().tolist() if 'CHV_CTE' in df_cte.columns else []

                # Combine
                all_keys = [{'CHAVE': k, 'TIPO': 'NFE'} for k in keys_nfe if k] + \
                           [{'CHAVE': k, 'TIPO': 'CTE'} for k in keys_cte if k]
                df_keys = pd.DataFrame(all_keys)

            else: # XML
                # processar_pasta_xml returns: df_totais, df_itens, df_cte_xml
                df_nfe, _, df_cte = processar_pasta_xml(Path(self.input_path_val))

                keys_nfe = df_nfe['CHV_NFE'].unique().tolist() if 'CHV_NFE' in df_nfe.columns else []
                keys_cte = df_cte['CHV_CTE'].unique().tolist() if 'CHV_CTE' in df_cte.columns else []

                all_keys = [{'CHAVE': k, 'TIPO': 'NFE'} for k in keys_nfe if k] + \
                           [{'CHAVE': k, 'TIPO': 'CTE'} for k in keys_cte if k]
                df_keys = pd.DataFrame(all_keys)

            if not df_keys.empty:
                output_file = f"Chaves_Extraidas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                df_keys.to_excel(output_file, index=False)
                self.status_text.value = f"Sucesso! Salvo em: {output_file}"
                self.status_text.color = "green"
            else:
                self.status_text.value = "Nenhuma chave encontrada."
                self.status_text.color = "orange"

        except Exception as ex:
            self.status_text.value = f"Erro: {ex}"
            self.status_text.color = "red"

        self.extract_btn.disabled = False
        self.progress_bar.visible = False
        self.update()
