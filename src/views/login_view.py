import flet as ft
from src.utils.database import get_db_connection
import bcrypt

class LoginView(ft.Container):
    def __init__(self, page, on_login_success):
        super().__init__()
        self.page = page
        self.on_login_success = on_login_success
        self.expand = True
        self.alignment = ft.alignment.center

        self.username_field = ft.TextField(label="Usuário", width=300)
        self.password_field = ft.TextField(label="Senha", password=True, can_reveal_password=True, width=300)
        self.error_text = ft.Text(color="red")

        self.content = ft.Column(
            [
                ft.Text("Login", size=30, weight="bold"),
                self.username_field,
                self.password_field,
                self.error_text,
                ft.ElevatedButton("Entrar", on_click=self.login),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )

    def login(self, e):
        username = self.username_field.value
        password = self.password_field.value

        if not username or not password:
            self.error_text.value = "Preencha todos os campos."
            self.update()
            return

        # Database check
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username = ?", (username,))
        user_row = cursor.fetchone()
        conn.close()

        if user_row:
             db_pass = user_row['password']
             # db_pass is stored as string, encode to bytes for bcrypt
             if bcrypt.checkpw(password.encode('utf-8'), db_pass.encode('utf-8')):
                 user = {
                     "username": user_row['username'],
                     "is_admin": bool(user_row['is_admin']),
                     "permissions": user_row['permissions']
                 }
                 self.on_login_success(user)
                 return

        self.error_text.value = "Usuário ou senha incorretos."
        self.update()
