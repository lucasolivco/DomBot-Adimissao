import tkinter as tk
from tkinter import messagebox, scrolledtext
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import findwindows, timings
import win32gui
import win32con
import time
import logging
from datetime import datetime
import os
import traceback
import threading
from pywinauto.timings import wait_until, wait_until_passes

class AutomacaoGUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("DomBot - Folha de Ponto")
        self.window.geometry("820x680")
        self.window.protocol("WM_DELETE_WINDOW", self.ao_fechar)

        self.executando = False
        self.thread_automacao = None

        self.logs_dir = os.path.join(os.path.dirname(__file__), "logs")
        if not os.path.exists(self.logs_dir):
            os.makedirs(self.logs_dir)

        self.setup_file_logging()

        self.logger = logging.getLogger('AutomacaoDominio')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []

        class GUIHandler(logging.Handler):
            def __init__(self, gui):
                super().__init__()
                self.gui = gui
            def emit(self, record):
                msg = self.format(record)
                self.gui.adicionar_log(msg)
        self.gui_handler = GUIHandler(self)
        formatter = logging.Formatter('%(message)s')
        self.gui_handler.setFormatter(formatter)
        self.logger.addHandler(self.gui_handler)

        self.companies = []
        self.current_funcionarios = []

        self.criar_interface()

    def setup_file_logging(self):
        data_atual = datetime.now().strftime("%Y-%m-%d")
        self.success_logger = logging.getLogger('SuccessLog')
        self.success_logger.setLevel(logging.INFO)
        if not self.success_logger.handlers:
            success_handler = logging.FileHandler(
                os.path.join(self.logs_dir, f'success_{data_atual}.log'),
                encoding='utf-8'
            )
            success_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
            )
            self.success_logger.addHandler(success_handler)
        self.error_logger = logging.getLogger('ErrorLog')
        self.error_logger.setLevel(logging.ERROR)
        if not self.error_logger.handlers:
            error_handler = logging.FileHandler(
                os.path.join(self.logs_dir, f'error_{data_atual}.log'),
                encoding='utf-8'
            )
            error_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
            )
            self.error_logger.addHandler(error_handler)

    def criar_interface(self):
        main_frame = tk.Frame(self.window)
        main_frame.pack(fill="both", expand=True, padx=12, pady=12)

        cadastro_frame = tk.LabelFrame(main_frame, text="Cadastro", padx=8, pady=8)
        cadastro_frame.pack(fill="x", pady=4)

        empresa_frame = tk.Frame(cadastro_frame)
        empresa_frame.pack(fill="x", pady=2)
        tk.Label(empresa_frame, text="Código da Empresa:").pack(side="left")
        self.empresa_entry = tk.Entry(empresa_frame, width=13)
        self.empresa_entry.pack(side="left", padx=6)

        funcionario_frame = tk.Frame(cadastro_frame)
        funcionario_frame.pack(fill="x", pady=2)
        tk.Label(funcionario_frame, text="Código do Funcionário:").pack(side="left")
        self.funcionario_entry = tk.Entry(funcionario_frame, width=13)
        self.funcionario_entry.pack(side="left", padx=6)
        tk.Button(funcionario_frame, text="Adicionar Funcionário", command=self.adicionar_funcionario, width=18).pack(side="left", padx=10)

        self.funcionarios_textbox = scrolledtext.ScrolledText(cadastro_frame, height=5, width=70, state="disabled")
        self.funcionarios_textbox.pack(fill="x", pady=3)
        self.funcionarios_textbox.config(state="normal")
        self.funcionarios_textbox.insert("end", "Funcionários da empresa atual:\n")
        self.funcionarios_textbox.config(state="disabled")

        tk.Button(cadastro_frame, text="Adicionar Empresa", command=self.adicionar_empresa, width=18).pack(pady=3)

        separator = tk.Label(main_frame, text="-" * 80)
        separator.pack(pady=3)

        control_frame = tk.Frame(main_frame)
        control_frame.pack(fill="x", pady=2)
        self.btn_iniciar = tk.Button(control_frame, text="Iniciar Automação", command=self.iniciar_automacao_thread, height=2, width=24)
        self.btn_iniciar.pack(side="left", expand=True, fill="x", padx=4)
        self.btn_parar = tk.Button(control_frame, text="Parar", command=self.parar_automacao, height=2, width=8, state="disabled")
        self.btn_parar.pack(side="left", expand=True, fill="x", padx=4)

        progress_frame = tk.Frame(main_frame)
        progress_frame.pack(fill="x", pady=6)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = tk.Scale(progress_frame, variable=self.progress_var, orient="horizontal", from_=0, to=100, showvalue=False, length=600, state="disabled")
        self.progress_bar.pack(fill="x", padx=10)
        self.status_var = tk.StringVar(value="Aguardando início...")
        tk.Label(progress_frame, textvariable=self.status_var, height=2).pack()

        log_frame = tk.LabelFrame(main_frame, text="Logs", padx=8, pady=8)
        log_frame.pack(fill="both", expand=True, pady=7)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=14, width=90)
        self.log_text.pack(fill="both", expand=True)
        tk.Button(log_frame, text="Limpar Logs", command=self.limpar_logs, width=18).pack(pady=3)

    def adicionar_funcionario(self):
        funcionario = self.funcionario_entry.get().strip()
        if funcionario:
            self.current_funcionarios.append(funcionario)
            self.funcionarios_textbox.config(state="normal")
            self.funcionarios_textbox.insert("end", f"{funcionario}\n")
            self.funcionarios_textbox.see("end")
            self.funcionarios_textbox.config(state="disabled")
            self.funcionario_entry.delete(0, "end")
            self.adicionar_log(f"Funcionário {funcionario} adicionado para a empresa atual")
        else:
            self.adicionar_log("Informe o código do funcionário antes de adicionar")

    def adicionar_empresa(self):
        empresa = self.empresa_entry.get().strip()
        if not empresa:
            self.adicionar_log("Informe o código da empresa")
            return
        if not self.current_funcionarios:
            self.adicionar_log("Cadastre ao menos um funcionário para a empresa")
            return
        self.companies.append({"empresa": empresa, "funcionarios": self.current_funcionarios.copy()})
        self.adicionar_log(f"Empresa {empresa} com {len(self.current_funcionarios)} funcionário(s) adicionada.")
        self.empresa_entry.delete(0, "end")
        self.funcionario_entry.delete(0, "end")
        self.funcionarios_textbox.config(state="normal")
        self.funcionarios_textbox.delete("1.0", "end")
        self.funcionarios_textbox.insert("end", "Funcionários da empresa atual:\n")
        self.funcionarios_textbox.config(state="disabled")
        self.current_funcionarios = []

    def limpar_logs(self):
        self.log_text.delete("1.0", "end")
        self.adicionar_log("Log limpo")

    def adicionar_log(self, mensagem):
        self.log_text.insert("end", f"{datetime.now().strftime('%H:%M:%S')} - {mensagem}\n")
        self.log_text.see("end")
        self.window.update_idletasks()

    def atualizar_progresso(self, atual, total):
        porcentagem = (atual / total) * 100 if total > 0 else 0
        self.progress_var.set(porcentagem)
        self.status_var.set(f"Processando: {atual}/{total} ({porcentagem:.1f}%)")
        self.window.update_idletasks()

    def iniciar_automacao_thread(self):
        if self.executando:
            self.adicionar_log("Automação já em execução")
            return
        if not self.companies:
            self.adicionar_log("Nenhuma empresa foi cadastrada. Cadastre pelo menos uma empresa com seus funcionários.")
            return
        self.thread_automacao = threading.Thread(target=self.iniciar_automacao)
        self.thread_automacao.daemon = True
        self.thread_automacao.start()
        self.btn_iniciar.config(state="disabled")
        self.btn_parar.config(state="normal")
        self.progress_bar.config(state="normal")

    def parar_automacao(self):
        if self.executando:
            self.executando = False
            self.adicionar_log("Solicitação de parada enviada. Aguardando conclusão...")
            self.status_var.set("Interrompendo...")

    def ao_fechar(self):
        if self.executando:
            if messagebox.askyesno("Confirmação", "Existe uma automação em execução. Deseja realmente sair?"):
                self.executando = False
                self.window.after(1000, self.window.destroy)
        else:
            self.window.destroy()

    def iniciar_automacao(self):
        self.adicionar_log("Iniciando automação...")
        self.status_var.set("Em execução...")
        self.executando = True

        total_iteracoes = sum(len(emp["funcionarios"]) for emp in self.companies)
        iteracao_atual = [0]

        try:
            automacao = DominioAutomation(self.logger, self)
            if not automacao.connect_to_dominio():
                erro_msg = "Erro: Não foi possível conectar ao Domínio"
                self.adicionar_log(erro_msg)
                self.error_logger.error(erro_msg)
                return

            for comp in self.companies:
                empresa = comp["empresa"]
                funcionarios = comp["funcionarios"]

                if not self.executando:
                    break

                if not automacao.switch_to_company(empresa):
                    self.error_logger.error(f"Falha ao trocar para a empresa {empresa}")
                    self.adicionar_log(f"Falha ao trocar para a empresa {empresa}")
                    self.executando = False
                    break

                success = automacao.processar_funcionarios_empresa(
                    empresa, funcionarios,
                    progresso_callback=lambda: self.atualizar_progresso(iteracao_atual[0], total_iteracoes),
                    sucesso_callback=lambda funcionario: self.success_logger.info(f"Empresa {empresa} - Funcionário {funcionario} - Enviado com sucesso"),
                    erro_callback=lambda funcionario: self.error_logger.error(f"Empresa {empresa} - Funcionário {funcionario} - Erro no envio"),
                    loop_control=lambda: self.executando,
                    avance_callback=lambda: (iteracao_atual.__setitem__(0, iteracao_atual[0]+1))
                )

                if not success:
                    self.adicionar_log(f"Processo interrompido na empresa {empresa}")
                    break

            self.status_var.set("Processamento concluído")
            self.progress_var.set(100)
        except Exception as e:
            erro_msg = f"Erro crítico: {str(e)}"
            self.error_logger.error(erro_msg)
            self.adicionar_log(erro_msg)
            self.adicionar_log(traceback.format_exc())
            self.status_var.set("Erro no processamento")
        finally:
            self.executando = False
            self.btn_iniciar.config(state="normal")
            self.btn_parar.config(state="disabled")
            self.progress_bar.config(state="disabled")

    def executar(self):
        self.window.mainloop()


class DominioAutomation:
    def __init__(self, logger, gui):
        timings.Timings.window_find_timeout = 20
        self.app = None
        self.main_window = None
        self.logger = logger
        self.gui = gui

    def log(self, message):
        self.logger.info(message)

    def find_dominio_window(self):
        try:
            windows = findwindows.find_windows(title_re=".*Domínio Folha.*")
            if windows:
                return windows[0]
            self.log("Nenhuma janela do Domínio Folha encontrada.")
            return None
        except Exception as e:
            self.log(f"Erro ao procurar a janela do Domínio Folha: {str(e)}")
            return None

    def connect_to_dominio(self):
        try:
            handle = self.find_dominio_window()
            if not handle:
                self.log("Não foi possível encontrar a janela do Domínio Folha.")
                return False
            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)
            win32gui.SetForegroundWindow(handle)
            time.sleep(0.5)
            self.app = Application(backend="uia").connect(handle=handle)
            self.main_window = self.app.window(handle=handle)
            return True
        except Exception as e:
            self.log(f"Erro ao conectar ao Domínio Folha: {str(e)}")
            return False

    def wait_for_window(self, titulo, timeout=30):
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                window = self.app.window(title=titulo)
                if window.exists():
                    return window
            except Exception:
                pass
            time.sleep(0.5)
        raise TimeoutError(f"Timeout esperando pela janela: {titulo}")

    def wait_and_check_window_closed(self, window, window_title, timeout=30):
        start_time = time.time()
        while time.time() - start_time < timeout:
            if not window.exists() or not window.is_visible():
                self.log(f"Janela '{window_title}' fechada com sucesso")
                return True
            else:
                self.log(f"Janela '{window_title}' não foi fechada")
                return False
            time.sleep(0.5)
        self.log(f"Aviso: Tempo máximo de espera atingido para fechamento da janela '{window_title}'")
        return False

    def fechar_janelas_filhas(self):
        try:
            self.log("Fechando janelas filhas antes de trocar de empresa...")
            janela_principal = self.main_window or self.app.top_window()
            filhos = janela_principal.children()
            for filho in filhos:
                try:
                    filho.close()
                    self.log(f"Janela '{filho.window_text()}' fechada.")
                except Exception as e:
                    self.log(f"Erro ao tentar fechar '{filho.window_text()}': {e}")
        except Exception as e:
            self.log(f"Erro ao fechar janelas filhas: {str(e)}")

    def switch_to_company(self, empresa):
        try:
            self.log(f"Trocando para a empresa {empresa}")
            handle = self.find_dominio_window()
            if not handle:
                self.log("Não foi possível encontrar a janela do Domínio Folha.")
                return False
            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)
            win32gui.SetForegroundWindow(handle)
            time.sleep(0.5)
            app = Application(backend="uia").connect(handle=handle)
            main_window = app.window(handle=handle)
            main_window.set_focus()
            time.sleep(0.5)

            send_keys('{F8}')
            time.sleep(1.5)

            troca_empresas_window = None
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    troca_empresas_window = main_window.child_window(
                        title="Troca de empresas",
                        class_name="FNWND3190"
                    )
                    if troca_empresas_window.exists():
                        break
                    troca_empresas_windows = main_window.children(title="Troca de empresas")
                    if troca_empresas_windows:
                        troca_empresas_window = troca_empresas_windows[0]
                        break
                except Exception:
                    if attempt == max_attempts - 1:
                        self.log("Janela 'Troca de empresas' não encontrada após várias tentativas.")
                        return False
                    time.sleep(1)
            if not troca_empresas_window:
                self.log("Janela 'Troca de empresas' não encontrada.")
                return False

            self.log("Janela 'Troca de empresas' visível")
            send_keys(str(empresa))
            time.sleep(1)
            send_keys('{ENTER}')
            time.sleep(6)
            self.wait_and_check_window_closed(troca_empresas_window, "Troca de empresas")

            try:
                aviso_window = main_window.child_window(
                    title="Avisos de Vencimento",
                    class_name="FNWND3190"
                )
                if aviso_window.exists() and aviso_window.is_visible():
                    self.log("Fechando 'Avisos de Vencimento'")
                    aviso_window.set_focus()
                    send_keys('{ESC}')
                    time.sleep(1)
                    send_keys('{ESC}')
            except Exception:
                self.log("Nenhuma janela de 'Avisos de Vencimento' encontrada")

            return True
        except Exception as e:
            self.log(f"Erro ao trocar para a empresa {empresa}: {str(e)}")
            return False

    def publicar_documento(self, nome_arquivo, empresa, auto_close=True):
        try:
            pub_window = self.main_window.child_window(title="Publicação de Documentos Externos")
            pub_window.wait("visible", timeout=30)

            onvio_processos = pub_window.child_window(class_name="Edit", auto_id="1013")
            campo_numero = pub_window.child_window(class_name="PBEDIT190", auto_id="1001")
            botao = pub_window.child_window(class_name="Button", auto_id="1003")

            onvio_processos.set_text(f"C:\\Documentos\\{nome_arquivo}.docx")
            campo_numero.set_text(empresa)

            pasta = pub_window.child_window(auto_id="1001", class_name="ComboBox")
            pasta.click_input()
            send_keys("pessoal/adimissionais")
            send_keys("{ENTER}")

            # botao.click_input()
            # self.log("Botão publicar clicado com sucesso!")

            try:
                atencao_window = self.wait_for_window("Atenção", timeout=10)
                atencao_window.set_focus()
                ok_button = atencao_window.child_window(title="OK")
                if ok_button.exists():
                    ok_button.click()
            except Exception:
                pass

            if auto_close:
                send_keys('{ESC}')
                time.sleep(1)

            return True
        except Exception as e:
            self.log(f"Erro na publicação: {str(e)}")
            return False

    def processar_funcionarios_empresa(self, empresa, funcionarios, progresso_callback=None, sucesso_callback=None, erro_callback=None, loop_control=lambda: True, avance_callback=None):
        """
        Processa todos os funcionários de uma empresa SEM trocar de empresa até acabar.
        Mantém janela de relatório aberta, volta para Admissionais RH Canella(1) após cada funcionário.
        """
        try:
            handle = self.find_dominio_window()
            if not handle:
                return False
            if win32gui.IsIconic(handle):
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                time.sleep(1)
            win32gui.SetForegroundWindow(handle)
            time.sleep(0.5)
            app = Application(backend="uia").connect(handle=handle)
            main_window = app.window(handle=handle)
            main_window.set_focus()
            time.sleep(0.5)

            # Abre menu Relatórios > Personalizados > Admissionais RH Canella(1)
            send_keys('%r')
            time.sleep(0.5)
            send_keys('i')
            time.sleep(0.5)
            send_keys('i')
            time.sleep(0.5)
            send_keys('{ENTER}')
            time.sleep(1)

            # Espera abrir gerenciador de relatórios
            relatorio_window = None
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    relatorio_window = main_window.child_window(
                        title="Gerenciador de Relatórios",
                        class_name="FNWND3190"
                    )
                    if relatorio_window.exists():
                        break
                except Exception:
                    if attempt == max_attempts - 1:
                        return False
                    time.sleep(1)
            if not relatorio_window or not relatorio_window.exists():
                return False

            rel_app = Application(backend='uia').connect(handle=relatorio_window.handle)
            tree = rel_app.window(class_name="FNWND3190").child_window(class_name="PBTreeView32_100")

            # Personalizados > Admissionais RH Canella(1)
            Personalizados = tree.child_window(title="Personalizados")
            if not Personalizados.exists():
                return False
            Personalizados.set_focus()
            Personalizados.click_input()
            time.sleep(0.5)
            Personalizados.click_input(double=True)
            time.sleep(1)
            Admissionais = tree.child_window(title="Admissionais RH Canella(1)")
            if not Admissionais.exists():
                return False
            Admissionais.set_focus()
            Admissionais.click_input(double=True)
            time.sleep(1)

            # Processa cada funcionário SEM fechar janela
            for idx, funcionario in enumerate(funcionarios):
                if not loop_control():
                    self.log("Automação interrompida pelo usuário")
                    return False
                self.log(f"Processando: Empresa {empresa} - Funcionário {funcionario}")
                # Preenche código funcionário
                send_keys('{TAB}' + str(funcionario))
                time.sleep(0.3)
                button_executar = relatorio_window.child_window(auto_id="1007", class_name="Button")
                button_executar.click_input()

                time.sleep(5)
                try:
                    app_word = wait_until_passes(
                        timeout=40,
                        retry_interval=2,
                        func=lambda: Application(backend="win32").connect(
                            title_re=".*Admissionais RH Canella.*",
                            class_name="OpusApp"
                        )
                    )
                except Exception:
                    self.log(f"Erro ao gerar relatório Word para funcionário {funcionario}")
                    if erro_callback:
                        erro_callback(funcionario)
                    return False

                send_keys('{F12}')
                time.sleep(2)

                try:
                    _ = wait_until_passes(
                        timeout=30,
                        retry_interval=2,
                        func=lambda: Application(backend="win32").connect(title_re="Salvar como").window(title_re="Salvar como")
                    )
                    nome_arquivo = f"{empresa} - funcionário {funcionario}"
                    send_keys(nome_arquivo)
                    time.sleep(1)
                    send_keys('{ENTER}')
                    time.sleep(3)
                    send_keys('%{F4}')
                    time.sleep(2)
                except Exception:
                    self.log(f"Erro ao salvar relatório Word para funcionário {funcionario}")
                    if erro_callback:
                        erro_callback(funcionario)
                    return False

                self.main_window.set_focus()
                time.sleep(1)

                # Publica documento
                try:
                    botao_publicar = self.main_window.child_window(
                        auto_id="picturePublicacaoDocumentosExternos",
                        class_name="WindowsForms10.Window.8.app.0.378734a"
                    )
                    botao_publicar.wait("visible", timeout=15)
                    botao_publicar.click_input()
                    time.sleep(2)
                except Exception:
                    self.log("Erro ao clicar no botão Publicar")
                    if erro_callback:
                        erro_callback(funcionario)
                    return False

                if not self.publicar_documento(nome_arquivo, empresa, auto_close=True):
                    self.log(f"Erro ao publicar documento para funcionário {funcionario}")
                    if erro_callback:
                        erro_callback(funcionario)
                    return False

                if sucesso_callback:
                    sucesso_callback(funcionario)
                if avance_callback:
                    avance_callback()
                if progresso_callback:
                    progresso_callback()

                # Se não for o último funcionário, volta para Admissionais RH Canella(1)
                if idx < len(funcionarios) - 1:
                    # Volta para o treeview e foca novamente
                    try:
                        tree = relatorio_window.child_window(class_name="PBTreeView32_100")
                        Personalizados = tree.child_window(title="Personalizados")
                        if Personalizados.exists():
                            Personalizados.set_focus()
                            Personalizados.click_input()
                            time.sleep(0.5)
                            Personalizados.click_input(double=True)
                            time.sleep(1)
                            Admissionais = tree.child_window(title="Admissionais RH Canella(1)")
                            if Admissionais.exists():
                                Admissionais.set_focus()
                                Admissionais.click_input(double=True)
                                time.sleep(1)
                            else:
                                self.log("Não encontrou Admissionais RH Canella(1) ao retornar")
                                return False
                        else:
                            self.log("Não encontrou Personalizados ao retornar")
                            return False
                    except Exception:
                        self.log("Erro ao tentar reposicionar no relatório após funcionário")
                        return False

                    self.main_window.set_focus()
                    time.sleep(1)
                # Se for o último funcionário, fecha janelas filhas e retorna
                else:
                    self.fechar_janelas_filhas()
                    self.main_window.set_focus()
                    send_keys('{ESC}')
                    time.sleep(2)
            return True
        except Exception as e:
            self.log(f"Erro ao processar funcionários da empresa {empresa}: {str(e)}")
            self.log(traceback.format_exc())
            return False

def main():
    gui = AutomacaoGUI()
    gui.executar()

if __name__ == "__main__":
    main()