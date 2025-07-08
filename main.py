import customtkinter as ctk
import threading
import schedule
import time
import os
import sys
from contextlib import redirect_stdout
from io import StringIO

# Adicionar suporte ao system tray
try:
    import pystray
    from PIL import Image
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False
    print("pystray não disponível. Instale com: pip install pystray")

from auto_103 import main as run_automation_logic

# --- Configurações de Aparência ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Classe auxiliar para redirecionar o 'print' para a GUI
class GuiLogger(StringIO):
    """
    Uma classe que herda de StringIO para capturar a saída do 'print' 
    e enviá-la para a nossa função de log.
    """
    def __init__(self, log_function):
        super().__init__()
        self.log_function = log_function

    def write(self, text):
        # Chama o método write da classe pai
        super().write(text)
        # A saída do print pode vir com quebras de linha extras.
        # Nós só queremos registrar o texto se ele não estiver vazio.
        if text.strip():
            self.log_function(text.strip())

class AutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação 103 - SSW")
        self.root.geometry("600x520")
        self.root.resizable(False, False)
        
        # Configurar para não aparecer na barra de tarefas
        self.root.attributes('-toolwindow', True)
        
        # Configurar para iniciar minimizado
        self.root.withdraw()  # Esconde a janela inicialmente
        
        # Configurar protocolo de fechamento para esconder em vez de fechar
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        
        self.scheduler_running = False
        self.tray_icon = None
        self.create_widgets()
        self.setup_tray()
        
        # Mostrar notificação de inicialização
        self.show_notification("Automação 103", "Aplicativo iniciado no system tray")

    def create_widgets(self):
        main_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)

        ctk.CTkLabel(main_frame, text="Automação 103 - Situações Coletas", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(0, 20))

        self.run_now_button = ctk.CTkButton(
            main_frame, text="Executar Agora", command=self.start_automation_thread,
            font=ctk.CTkFont(size=13, weight="bold"), height=40, corner_radius=8
        )
        self.run_now_button.pack(fill="x", pady=(0, 10))

        schedule_frame = ctk.CTkFrame(main_frame, corner_radius=8)
        schedule_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(schedule_frame, text="Agendamento Diário", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=10)

        time_frame = ctk.CTkFrame(schedule_frame, fg_color="transparent")
        time_frame.pack(pady=(0, 15), padx=10, fill="x")
        time_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkLabel(time_frame, text="Hora:").grid(row=0, column=0, padx=(0,5), sticky="e")
        self.hour_entry = ctk.CTkEntry(time_frame, width=60)
        self.hour_entry.grid(row=0, column=1, sticky="w")
        self.hour_entry.insert(0, "08")

        ctk.CTkLabel(time_frame, text="Minuto:").grid(row=0, column=2, padx=(10,5), sticky="e")
        self.minute_entry = ctk.CTkEntry(time_frame, width=60)
        self.minute_entry.grid(row=0, column=3, sticky="w")
        self.minute_entry.insert(0, "30")

        schedule_buttons_frame = ctk.CTkFrame(schedule_frame, fg_color="transparent")
        schedule_buttons_frame.pack(pady=(0, 15), padx=20, fill="x")
        schedule_buttons_frame.grid_columnconfigure(0, weight=1)
        schedule_buttons_frame.grid_columnconfigure(1, weight=1)

        self.schedule_button = ctk.CTkButton(
            schedule_buttons_frame, text="Agendar Execução", command=self.schedule_automation,
            fg_color="#2E8B57", hover_color="#3CB371"
        )
        self.schedule_button.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        self.cancel_button = ctk.CTkButton(
            schedule_buttons_frame, text="Cancelar Agendamento", command=self.cancel_schedule,
            fg_color="#8B0000", hover_color="#B22222", state="disabled"
        )
        self.cancel_button.grid(row=0, column=1, padx=(5, 0), sticky="ew")

        ctk.CTkLabel(main_frame, text="Status da Automação:", font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(10, 5))
        self.log_area = ctk.CTkTextbox(main_frame, state="disabled", corner_radius=8, font=('Consolas', 11))
        self.log_area.pack(fill="both", expand=True)

    def setup_tray(self):
        if not TRAY_AVAILABLE:
            return
            
        # Criar ícone para o system tray
        try:
            # Tentar usar o ícone existente
            icon_path = "icon.png"
            if os.path.exists(icon_path):
                image = Image.open(icon_path)
            else:
                # Criar um ícone simples se não existir
                image = Image.new('RGB', (64, 64), color='blue')
        except:
            # Fallback para ícone simples
            image = Image.new('RGB', (64, 64), color='blue')
        
        # Criar menu do system tray
        menu = pystray.Menu(
            pystray.MenuItem("Mostrar Automação", self.show_window),
            pystray.MenuItem("Sair", self.quit_app)
        )
        
        self.tray_icon = pystray.Icon("auto_103", image, "Automação 103", menu)
        
        # Iniciar o system tray em uma thread separada
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_window(self):
        """Mostra a janela principal"""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        # Garantir que a janela não apareça na barra de tarefas
        self.root.attributes('-toolwindow', True)

    def hide_window(self):
        """Esconde a janela principal"""
        self.root.withdraw()
        self.show_notification("Assistente de Automação", "Aplicativo continua rodando em segundo plano")

    def show_notification(self, title, message):
        """Mostra uma notificação no system tray"""
        if self.tray_icon:
            self.tray_icon.notify(title, message)



    def quit_app(self):
        """Fecha o aplicativo completamente"""
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.quit()

    def log_message(self, message):
        self.log_area.configure(state="normal")
        self.log_area.insert("end", f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.log_area.configure(state="disabled")
        self.log_area.see("end")
        
        # Mostrar notificação no system tray para mensagens importantes
        if "ERRO" in message.upper() or "concluída" in message.lower():
            self.show_notification("Assistente de Automação", message)

    def _update_buttons_state(self, is_running):
        if is_running:
            self.run_now_button.configure(state="disabled")
            self.schedule_button.configure(state="disabled")
            self.cancel_button.configure(state="disabled")
        else:
            self.run_now_button.configure(state="normal")
            self.schedule_button.configure(state="normal")
            if schedule.jobs:
                self.cancel_button.configure(state="normal")
            else:
                self.cancel_button.configure(state="disabled")

    def cancel_schedule(self):
        if schedule.jobs:
            schedule.clear()
            self.log_message("Agendamento cancelado com sucesso.")
            self.cancel_button.configure(state="disabled")
            self.show_notification("Assistente de Automação", "Agendamento cancelado")
        else:
            self.log_message("Não há nenhum agendamento ativo para cancelar.")

    def start_automation_thread(self):
        self.log_message("Iniciando a automação...")
        self.show_notification("Assistente de Automação", "Iniciando automação...")
        self._update_buttons_state(is_running=True)
        threading.Thread(target=self.run_automation_wrapper, daemon=True).start()

    def run_automation_wrapper(self):
        """
        Executa a automação e redireciona a saída do 'print' para a GUI
        de uma forma segura e moderna.
        """
        # Cria uma instância do nosso logger, passando a função de log da GUI
        gui_logger = GuiLogger(self.log_message)

        try:
            # O 'with' garante que o redirecionamento aconteça apenas aqui dentro
            # e seja desfeito automaticamente no final, mesmo se der erro.
            with redirect_stdout(gui_logger):
                run_automation_logic() # Chama a sua função principal
                print("Automação concluída com sucesso!")

        except Exception as e:
            # Como ainda estamos dentro do 'with', este print também será redirecionado
            print(f"ERRO: Ocorreu uma falha na automação: {e}")
        finally:
            # Atualiza o estado dos botões na thread principal da GUI
            self.root.after(0, self._update_buttons_state, False)

    def schedule_automation(self):
        hour = self.hour_entry.get()
        minute = self.minute_entry.get()
        if not hour.isdigit() or not minute.isdigit() or not (0 <= int(hour) <= 23 and 0 <= int(minute) <= 59):
            self.log_message("Erro: Por favor, insira um horário válido (HH:MM).")
            return
            
        schedule_time = f"{int(hour):02}:{int(minute):02}"
        self.log_message(f"Automação agendada para rodar diariamente às {schedule_time}.")
        self.show_notification("Assistente de Automação", f"Agendado para {schedule_time}")
        
        schedule.clear()
        schedule.every().day.at(schedule_time).do(self.start_automation_thread)
        
        self.cancel_button.configure(state="normal")
        
        if not self.scheduler_running:
            self.scheduler_running = True
            threading.Thread(target=self.run_scheduler, daemon=True).start()
            self.log_message("Serviço de agendamento iniciado.")

    def run_scheduler(self):
        while self.scheduler_running:
            schedule.run_pending()
            time.sleep(1)

if __name__ == "__main__":
    root = ctk.CTk()
    app = AutomationApp(root)
    root.mainloop()