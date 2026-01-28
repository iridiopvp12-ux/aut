import flet as ft
from src.utils.database import get_db_connection

class DashboardView(ft.Container):
    def __init__(self):
        super().__init__()
        self.expand = True
        self.padding = 20
        self.content = ft.Column(
            [
                ft.Text("Dashboard", size=30, weight="bold"),
                ft.Text("Bem-vindo ao SiegAuto - Sistema de Automação Fiscal.", size=16),
                ft.Divider(),

                ft.Row(
                    [
                        self.create_stat_card("Usuários", self.get_user_count(), ft.Icons.PEOPLE, ft.Colors.BLUE),
                        self.create_stat_card("Status do Sistema", "Online", ft.Icons.CHECK_CIRCLE, ft.Colors.GREEN),
                        # Future: Add stats for analyses run
                    ],
                    alignment=ft.MainAxisAlignment.START,
                    spacing=20
                ),

                ft.Divider(),
                ft.Text("Acesso Rápido", size=20, weight="bold"),
                ft.Row(
                    [
                        self.create_action_card("Analisador Fiscal (SPED)", "Conciliação Fiscal e Apuração", ft.Icons.DESCRIPTION, ft.Colors.ORANGE, "SPED"),
                        self.create_action_card("Invest & Contribuições", "Incentivos e PIS/COFINS", ft.Icons.MONETIZATION_ON, ft.Colors.PURPLE, "Invest"),
                    ],
                    alignment=ft.MainAxisAlignment.START,
                    spacing=20
                )
            ]
        )

    def create_stat_card(self, title, value, icon, color):
        return ft.Container(
            content=ft.Column(
                [
                    ft.Icon(icon, color=color, size=30),
                    ft.Text(str(value), size=24, weight="bold"),
                    ft.Text(title, size=14, color=ft.Colors.GREY_500),
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            width=200,
            height=120,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=10,
            padding=10,
            bgcolor=ft.Colors.WHITE10,
        )

    def create_action_card(self, title, subtitle, icon, color, nav_target):
        return ft.Container(
            content=ft.Row(
                [
                    ft.Icon(icon, color=color, size=40),
                    ft.Column(
                        [
                            ft.Text(title, size=16, weight="bold"),
                            ft.Text(subtitle, size=12, color=ft.Colors.GREY_500),
                        ],
                        alignment=ft.MainAxisAlignment.CENTER,
                    )
                ],
                alignment=ft.MainAxisAlignment.START,
            ),
            width=300,
            height=100,
            border=ft.border.all(1, ft.Colors.GREY_300),
            border_radius=10,
            padding=15,
            bgcolor=ft.Colors.WHITE10,
            on_click=lambda e: print(f"Navigate to {nav_target}") # Navigation handled by Rail usually, this is visual for now
        )

    def get_user_count(self):
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM users")
            count = cursor.fetchone()[0]
            conn.close()
            return count
        except:
            return "?"
