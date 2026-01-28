import flet as ft
from src.views.login_view import LoginView
from src.views.dashboard_view import DashboardView
from src.views.admin_view import AdminView
from src.views.sped_view import SpedView
from src.views.settings_view import SettingsView
from src.utils.database import initialize_db
from src.utils.logger import log_action

# REMOVIDO 'async' AQUI
def main(page: ft.Page):
    # --- Configurações da Janela ---
    page.title = "SiegAuto - Sistema Contabilidade"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 1200
    page.window_height = 800
    page.padding = 0
    
    # Define o diretório de assets e o ícone da janela
    page.assets_dir = "assets" 
    page.window_icon = "logo.png" 

    # --- Inicializa Banco de Dados ---
    initialize_db()

    # --- Estado da Aplicação ---
    current_user = None

    # --- Elementos de UI ---
    rail = ft.NavigationRail(
        selected_index=0,
        label_type=ft.NavigationRailLabelType.ALL,
        min_width=100,
        min_extended_width=400,
        group_alignment=-0.9,
        destinations=[],
        on_change=None 
    )

    page_content = ft.Container(expand=True, padding=20)

    # --- Funções de Navegação e Lógica ---

    def logout(e):
        nonlocal current_user
        if current_user:
            log_action(f"User logged out: {current_user['username']}")
        
        current_user = None
        page.clean()
        page.add(LoginView(page, on_login_success))
        page.update()

    def get_destination(icon, label, selected_icon=None):
        return ft.NavigationRailDestination(
            icon=icon,
            selected_icon=selected_icon,
            label=label
        )

    def on_nav_change(e):
        if not rail.destinations:
            return

        index = e.control.selected_index
        selected_label = rail.destinations[index].label

        # Limpa o conteúdo atual
        page_content.content = None

        # Roteamento das Telas
        if selected_label == "Dashboard":
            page_content.content = DashboardView()
        elif selected_label == "Admin":
            page_content.content = AdminView(page)
        elif selected_label == "SPED":
            page_content.content = SpedView(page)
        elif selected_label == "Configurações":
            page_content.content = SettingsView(page)

        page_content.update()

    rail.on_change = on_nav_change

    def on_login_success(user):
        nonlocal current_user
        current_user = user
        log_action(f"User logged in: {user['username']}")

        # --- Lógica de Permissões ---
        dests = []
        permissions = user.get('permissions', '').split(',')
        is_admin = user.get('is_admin', False)

        def has_perm(p):
            return is_admin or p in permissions or "all" in permissions

        # Construção dinâmica do Menu
        if has_perm("dashboard"):
            dests.append(get_destination(ft.Icons.DASHBOARD_OUTLINED, "Dashboard", ft.Icons.DASHBOARD))
        
        if is_admin:
            dests.append(get_destination(ft.Icons.ADMIN_PANEL_SETTINGS_OUTLINED, "Admin", ft.Icons.ADMIN_PANEL_SETTINGS))

        if has_perm("sped"):
            dests.append(get_destination(ft.Icons.DESCRIPTION_OUTLINED, "SPED", ft.Icons.DESCRIPTION))

        if has_perm("settings"):
            dests.append(get_destination(ft.Icons.SETTINGS_OUTLINED, "Configurações", ft.Icons.SETTINGS))

        rail.destinations = dests
        rail.trailing = ft.IconButton(ft.Icons.LOGOUT, on_click=logout, tooltip="Sair")

        # Carregamento da tela inicial após login
        if dests:
            rail.selected_index = 0
            first_label = dests[0].label
            
            if first_label == "Dashboard":
                page_content.content = DashboardView()
            elif first_label == "Admin":
                page_content.content = AdminView(page)
            elif first_label == "SPED":
                page_content.content = SpedView(page)
            elif first_label == "Configurações":
                page_content.content = SettingsView(page)
        else:
            page_content.content = ft.Text("Sem permissões de acesso.", size=20, color="red")

        # Monta o Layout Principal
        page.clean()
        page.add(
            ft.Row(
                [
                    rail,
                    ft.VerticalDivider(width=1),
                    page_content
                ],
                expand=True
            )
        )
        page.update()

    # --- Início do App ---
    # Adiciona a view de login ao iniciar
    page.add(LoginView(page, on_login_success))
    page.update()

if __name__ == "__main__":
    ft.app(target=main)