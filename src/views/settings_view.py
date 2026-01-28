import flet as ft

class SettingsView(ft.Container):
    def __init__(self, page):
        super().__init__()
        self.page = page
        self.expand = True
        self.padding = 20
        self.content = ft.Column(
            [
                ft.Text("Configurações", size=30, weight="bold"),
                ft.Text("Preferências do Sistema (Em Breve)", size=20),
            ]
        )
