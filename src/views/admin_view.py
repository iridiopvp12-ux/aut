import flet as ft

class AdminView(ft.Container):
    def __init__(self, page):
        super().__init__()
        self.page = page
        self.expand = True
        self.padding = 20
        self.content = ft.Column(
            [
                ft.Text("Admin Panel", size=30, weight="bold"),
                ft.Text("Gerenciamento de Usuários (Em Breve)", size=20),
            ]
        )
