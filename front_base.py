"""
Front-end padrao Tkinter para RPAs.
Ler empresas do config.ini ([empresas]) e expor callbacks para RPA e Automacao.
"""

import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import configparser

INI_PATH = Path(__file__).with_name("config.ini")


class ConfigError(RuntimeError):
    """Erro simples para configuracao invalida."""


def carregar_empresas():
    cfg = configparser.ConfigParser()
    if not INI_PATH.exists():
        raise ConfigError(f"Arquivo de configuracao nao encontrado: {INI_PATH}")
    cfg.read(INI_PATH, encoding="utf-8")
    if "empresas" not in cfg or not cfg["empresas"]:
        raise ConfigError("Preencha [empresas] no config.ini (codigo = nome)")
    empresas = {}
    for _, nome in cfg["empresas"].items():
        empresas[nome] = nome  # usa o nome como display e código
    return empresas


def carregar_mes_ano_default():
    cfg = configparser.ConfigParser()
    if INI_PATH.exists():
        cfg.read(INI_PATH, encoding="utf-8")
    return cfg.get("geral", "mes_ano", fallback="11-2025")


class StatusWindow:
    """
    UI padrao para acompanhar execucao do RPA.
    Passe callbacks on_rpa(codigo, display) e on_automation(codigo, display).
    """

    def __init__(self, root, on_rpa, on_automation, titulo="Robo RPA"):
        self.root = root
        self.on_rpa = on_rpa
        self.on_automation = on_automation

        self.empresas = carregar_empresas()
        displays = list(self.empresas.keys())
        self.selected_empresa = tk.StringVar(value=displays[0])
        self.mes_ano_var = tk.StringVar(value=carregar_mes_ano_default())

        self.root.title(titulo)
        self.root.geometry("")
        self.root.minsize(700, 520)

        self.main_label = ttk.Label(root, text="Aguardando inicio...", font=("Segoe UI", 12, "bold"))
        self.main_label.pack(pady=(10, 5))

        ttk.Label(root, text="Selecionar Empresa:").pack(pady=(5, 0))
        self.empresa_selector = ttk.Combobox(
            root, values=displays, textvariable=self.selected_empresa, state="readonly", width=40, justify="center"
        )
        self.empresa_selector.pack(pady=(0, 10), anchor="center")

        ttk.Label(root, text="Mes/Ano (MM-AAAA):").pack(pady=(0, 0))
        self.mes_ano_entry = ttk.Entry(root, textvariable=self.mes_ano_var, width=10, justify="center")
        self.mes_ano_entry.pack(pady=(0, 10))

        ttk.Label(root, text="Progresso da Analise de Dados").pack(pady=(5, 0))
        self.overall_progress = ttk.Progressbar(root, orient="horizontal", length=600, mode="determinate")
        self.overall_progress.pack(padx=20)

        ttk.Label(root, text="Log de Atividades:").pack(pady=(10, 0))
        self.log_area = ScrolledText(root, wrap=tk.WORD, width=78, height=10)
        self.log_area.pack(pady=5, padx=10, fill="both", expand=False)
        # Permite copiar (Ctrl+C) e selecionar (Ctrl+A), bloqueia edicao
        self.log_area.bind("<Control-c>", self._copy_log)
        self.log_area.bind("<Control-a>", self._select_all_log)
        self.log_area.bind("<Key>", lambda e: "break")

        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(pady=(10, 10))

        self.start_rpa_button = ttk.Button(
            self.button_frame, text="Iniciar RPA", command=self.start_rpa, width=20
        )
        self.start_rpa_button.pack(side="left", padx=10)

        self.close_button = ttk.Button(
            self.button_frame, text="Fechar", command=self.root.destroy, state="disabled", width=20
        )
        self.close_button.pack(side="bottom", padx=10)

    # --- Helpers UI thread-safe -------------------------------------------------
    def _ui(self, func, *args, **kwargs):
        self.root.after(0, func, *args, **kwargs)

    def _lock_buttons(self):
        self._ui(lambda: self.start_rpa_button.config(state="disabled"))
        self._ui(lambda: self.close_button.config(state="disabled"))

    def _unlock_buttons(self):
        self._ui(lambda: self.start_rpa_button.config(state="normal"))
        self._ui(lambda: self.close_button.config(state="normal"))

    # --- Acoes de botao ---------------------------------------------------------
    def start_rpa(self):
        self._lock_buttons()
        display = self.empresa_selector.get()
        codigo = self.empresas[display]
        mes_ano = self.get_mes_ano()
        self.add_log_message(f"Iniciando RPA para: {display} ({mes_ano})")
        threading.Thread(target=self._wrap(self.on_rpa, codigo, display, mes_ano), daemon=True).start()

    def start_automation(self):
        self._lock_buttons()
        display = self.empresa_selector.get()
        codigo = self.empresas[display]
        mes_ano = self.get_mes_ano()
        self.add_log_message(f"Iniciando automacao para: {display} ({mes_ano})")
        threading.Thread(target=self._wrap(self.on_automation, codigo, display, mes_ano), daemon=True).start()

    def _wrap(self, target, codigo, display, mes_ano):
        def runner():
            try:
                target(codigo, display, mes_ano)
            except Exception as exc:
                self.add_log_message(f"[ERRO] {exc}")
            finally:
                self._unlock_buttons()
        return runner

    # --- API para atualizar UI --------------------------------------------------
    def update_main_label(self, message):
        self._ui(lambda: self.main_label.config(text=message))

    def update_progress(self, bar, value):
        self._ui(lambda: bar.config(value=value))

    def add_log_message(self, message):
        def append():
            self.log_area.insert(tk.END, f"{datetime.now():%H:%M:%S} - {message}\n")
            self.log_area.see(tk.END)
        self._ui(append)

    def finalize(self):
        self.update_main_label("Processo finalizado!")
        self._unlock_buttons()

    def enable_automation_button(self):
        self._ui(self.start_automation_button.config, state="normal")

    def get_mes_ano(self):
        return self.mes_ano_var.get().strip()

    def show_automation_alert(self):
        self._ui(
            messagebox.showinfo,
            "Automacao iniciada",
            "Clique em OK e mantenha o Dominio em foco. Nao use o computador ate concluir.",
        )

    def _copy_log(self, event=None):
        try:
            text = self.log_area.selection_get()
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except Exception:
            pass
        return "break"

    def _select_all_log(self, event=None):
        self.log_area.tag_add("sel", "1.0", "end")
        return "break"


def criar_janela(on_rpa, on_automation, titulo="Robo RPA"):
    root = tk.Tk()
    app = StatusWindow(root, on_rpa, on_automation, titulo=titulo)
    return root, app


if __name__ == "__main__":
    # Exemplo rapido
    def dummy_rpa(codigo, display):  # pragma: no cover - demo manual
        import time

        time.sleep(1)
        app.add_log_message(f"Rodou analise para {display} ({codigo})")
        app.finalize()

    def dummy_auto(codigo, display):  # pragma: no cover - demo manual
        import time

        time.sleep(1)
        app.add_log_message(f"Rodou automacao para {display} ({codigo})")
        app.finalize()

    root, app = criar_janela(dummy_rpa, dummy_auto, titulo="Demo RPA")
    root.mainloop()
