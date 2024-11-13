import tkinter as tk
from threading import Thread
from playwright.sync_api import sync_playwright
import itertools
import time
import subprocess

LOGIN_URL = 'https://canaime.com.br/sgp2rr/login/login_principal.php'  # URL de login do sistema Canaimé

class CanaimeLoginInterface:
    # noinspection PyInterpreter
    def __init__(self, root, test_mode=False):
        self.root = root
        self.test_mode = test_mode
        self.setup_window()
        self.create_widgets()

        # Variáveis para armazenar credenciais e controle de sessão
        self.username = None
        self.password = None
        self.browser = None
        self.page = None
        self.is_running = False

        # Evento para tratar o fechamento da janela
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_window(self):
        """Configura a janela principal da aplicação."""
        self.root.title("Login Canaimé")
        window_width, window_height = 300, 225
        self.center_window(window_width, window_height)
        self.root.attributes('-topmost', True)

    def center_window(self, width, height):
        """Centraliza a janela na tela."""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        pos_x = (screen_width - width) // 2
        pos_y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")

    def create_widgets(self):
        """Cria todos os widgets da interface."""
        self.username_label = tk.Label(self.root, text="Usuário:", anchor='w')
        self.username_label.pack(pady=(10, 2))

        self.username_entry = tk.Entry(self.root)
        self.username_entry.pack(pady=(0, 10))
        self.username_entry.focus_set()

        self.password_label = tk.Label(self.root, text="Senha:", anchor='w')
        self.password_label.pack(pady=(10, 2))

        self.password_entry = tk.Entry(self.root, show="*")
        self.password_entry.pack(pady=(0, 10))

        self.login_button = tk.Button(self.root, text="Login", command=self.start_login_process)
        self.login_button.pack(pady=10)

        # Label para mostrar status (como animação de carregamento)
        self.status_label = tk.Label(self.root, text="")
        self.status_label.pack(pady=10)

        # Vincular o evento de pressionar Enter ao método de login
        self.root.bind('<Return>', self.on_enter)

    def start_login_process(self):
        """Inicia o processo de login em uma thread separada."""
        self.login_button.config(state=tk.DISABLED)
        self.status_label.config(text="Realizando login...")
        self.is_running = True

        # Iniciar animação e processo de login
        Thread(target=self.loading_animation).start()
        Thread(target=self.execute_login).start()

    def loading_animation(self):
        """Anima a bolinha enquanto o login está em andamento."""
        for frame in itertools.cycle(["\u25d0", "\u25d3", "\u25d1", "\u25d2"]):
            if not self.is_running:
                break
            self.status_label.config(text=f"Realizando login... {frame}")
            time.sleep(0.2)

    def on_enter(self, event):
        """Método chamado quando a tecla Enter é pressionada."""
        self.start_login_process()

    def on_close(self):
        """Método chamado ao fechar a janela."""
        self.username = None
        self.password = None
        self.root.destroy()

    def execute_login(self):
        """Executa o login utilizando Playwright em uma thread separada."""
        username = self.username_entry.get()
        password = self.password_entry.get()

        if not username or not password:
            self.show_error("Usuário e senha são obrigatórios.")
            return

        try:
            with sync_playwright() as playwright:
                self.browser = playwright.chromium.launch(headless=not self.test_mode)
                context = self.browser.new_context()
                self.page = context.new_page()
                login_successful = self.perform_login(self.page, username, password)

                if login_successful:
                    if self.test_mode:
                        print(f"Usuário: {username}, Senha: {password}")

                    # Após o login, armazena as credenciais e fecha a interface
                    self.username = username
                    self.password = password
                    self.root.after(1000, self.root.destroy)
                else:
                    self.show_error("Usuário ou senha inválidos.")

        except Exception as e:
            self.show_error(e)

    def perform_login(self, page, username, password):
        """Realiza o processo de login utilizando Playwright."""
        page.goto(LOGIN_URL)
        page.fill("input[name='usuario']", username)
        page.fill("input[name='senha']", password)
        page.press("input[name='senha']", "Enter")
        page.wait_for_timeout(5000)

        # Verifica se o login foi bem-sucedido
        if page.locator('img').count() < 4:
            return False
        else:
            self.login_success()
            return True

    def login_success(self):
        """Atualiza a interface para mostrar sucesso no login."""
        self.is_running = False
        self.update_interface(lambda: self.status_label.config(text="Login efetuado com sucesso!"))

    def show_error(self, message):
        """Mostra mensagem de erro e habilita o botão de login novamente."""
        self.is_running = False
        self.update_interface(lambda: (
            self.status_label.config(text=message),
            self.login_button.config(state=tk.NORMAL)
        ))

    def update_interface(self, func):
        """Atualiza a interface da aplicação."""
        self.root.after(0, func)

    def get_credentials(self):
        """Retorna as credenciais de login."""
        return self.username, self.password


def run_canaime_login(test_mode=False):
    root = tk.Tk()
    app = CanaimeLoginInterface(root, test_mode=test_mode)
    root.mainloop()
    return app.get_credentials()


def Login(test=False):
    """Função para login sem JavaScript e com imagens desativadas."""
    username, password = run_canaime_login(test_mode=test)
    if not username or not password:
        return None  # Retorna None se o login não for bem-sucedido

    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=not test)
    context = browser.new_context(
        java_script_enabled=False,
        extra_http_headers={"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"}
    )
    context.route("**/*",
                  lambda route: route.abort() if route.request.resource_type == "image" else route.continue_())

    page = context.new_page()
    page.goto(LOGIN_URL)
    page.fill("input[name='usuario']", username)
    page.fill("input[name='senha']", password)
    page.press("input[name='senha']", "Enter")

    return page
