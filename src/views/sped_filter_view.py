import flet as ft
import pandas as pd
from pathlib import Path
from src.logic.sped_parser import extrair_dados_sped
import threading
from datetime import datetime

class SpedFilterView(ft.Container):
    def __init__(self, page: ft.Page):
        super().__init__()
        self.page = page
        self.expand = True
        self.padding = 20
        self.input_path_val = None

        # UI
        self.path_text = ft.Text("Nenhum SPED selecionado", italic=True)
        self.pick_file_dialog = ft.FilePicker(on_result=self.pick_file_result)
        self.page.overlay.append(self.pick_file_dialog)

        self.start_date_field = ft.TextField(label="Data Inicial (DDMMAAAA)", width=200)
        self.end_date_field = ft.TextField(label="Data Final (DDMMAAAA)", width=200)

        self.filter_btn = ft.ElevatedButton("Filtrar e Exportar", icon=ft.Icons.FILTER_ALT, on_click=self.start_filter, disabled=True)
        self.status_text = ft.Text("")
        self.progress_bar = ft.ProgressBar(width=400, visible=False)

        self.content = ft.Column([
            ft.Text("Filtro de SPED (Por Data)", size=30, weight="bold"),
            ft.Text("Exporta registros do SPED dentro do período selecionado para Excel.", size=14),
            ft.Divider(),
            ft.Row([
                ft.ElevatedButton("Selecionar SPED", icon=ft.Icons.UPLOAD_FILE, on_click=lambda _: self.pick_file_dialog.pick_files(allow_multiple=False, allowed_extensions=["txt"])),
                self.path_text
            ]),
            ft.Row([self.start_date_field, self.end_date_field]),
            ft.Divider(),
            self.filter_btn,
            self.progress_bar,
            self.status_text
        ])

    def pick_file_result(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.input_path_val = e.files[0].path
            self.path_text.value = e.files[0].name
            self.filter_btn.disabled = False
            self.update()

    def start_filter(self, e):
        s_date = self.start_date_field.value
        e_date = self.end_date_field.value

        if len(s_date) != 8 or len(e_date) != 8:
            self.status_text.value = "Formato de data inválido. Use DDMMAAAA."
            self.status_text.color = "red"
            self.update()
            return

        self.filter_btn.disabled = True
        self.progress_bar.visible = True
        self.status_text.value = "Filtrando..."
        self.update()

        t = threading.Thread(target=self.run_filter, args=(s_date, e_date), daemon=True)
        t.start()

    def run_filter(self, start_str, end_str):
        try:
            # Using sped_parser to get dataframes
            # NOTE: sped_parser currently doesn't return the DATE of the invoice in the main dataframe (C100).
            # I need to check sped_parser.py content.
            # Looking at src/logic/sped_parser.py:
            # It extracts: 'CHV_NFE', 'VL_DOC_SPED', 'ICMS_SPED', etc.
            # It DOES NOT extract DT_DOC (field 9 of C100) or DT_E_S (field 10).
            # Therefore I cannot filter by date using the EXISTING parser output without modifying it.

            # Alternative: Basic text parsing here for the filter.

            start_dt = datetime.strptime(start_str, "%d%m%Y")
            end_dt = datetime.strptime(end_str, "%d%m%Y")

            records = []

            with open(self.input_path_val, 'r', encoding='latin-1') as f:
                for line in f:
                    if not line.startswith('|'): continue
                    parts = line.split('|')
                    if len(parts) < 2: continue

                    reg = parts[1]

                    # Logic for C100, D100 date extraction
                    dt_val = None
                    if reg in ['C100', 'D100']:
                        # C100: |REG|IND_OPER|IND_EMIT|COD_PART|COD_MOD|COD_SIT|SER|NUM_DOC|CHV_NFE|DT_DOC|...
                        # Index:  1     2        3        4        5       6     7      8       9      10
                        if len(parts) > 10:
                            dt_val = parts[10] # DT_DOC usually

                    if dt_val and len(dt_val) == 8:
                        try:
                            row_dt = datetime.strptime(dt_val, "%d%m%Y")
                            if start_dt <= row_dt <= end_dt:
                                records.append(parts)
                        except:
                            pass

            if records:
                # Convert to simple dataframe just for export
                df = pd.DataFrame(records)
                output_file = f"SPED_Filtrado_{start_str}_{end_str}.xlsx"
                df.to_excel(output_file, index=False, header=False)
                self.status_text.value = f"Exportado {len(records)} registros para {output_file}"
                self.status_text.color = "green"
            else:
                self.status_text.value = "Nenhum registro encontrado no período."
                self.status_text.color = "orange"

        except Exception as ex:
            self.status_text.value = f"Erro: {ex}"
            self.status_text.color = "red"

        self.filter_btn.disabled = False
        self.progress_bar.visible = False
        self.update()
