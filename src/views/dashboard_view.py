import flet as ft

class DashboardView(ft.Container):
    def __init__(self):
        super().__init__()
        self.expand = True
        self.padding = 20
        self.content = ft.Column(
            [
                ft.Text("Dashboard", size=30, weight="bold"),
                ft.Text("Bem-vindo ao SiegAuto.", size=20),
            ]
        )
