"""
Front-end simples para conciliação: seletor de empresa, campo Mes/Ano, barra de progresso e popups.
Sem área de log; notificações via messagebox.
"""

import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
import configparser
from utils import resource_path

INI_PATH = Path(resource_path("config.ini"))


class ConfigError(RuntimeError):
    pass


def carregar_empresas():
    cfg = configparser.ConfigParser()
    cfg.optionxform = str  # preserva o nome exibido
    if not INI_PATH.exists():
        raise ConfigError(f"Arquivo de configuracao nao encontrado: {INI_PATH}")
    cfg.read(INI_PATH, encoding="utf-8")
    if "EMPRESAS" not in cfg or not cfg["EMPRESAS"]:
        raise ConfigError("Preencha [empresas] no config.ini (codigo = nome)")
    empresas = {}
    for nome_empresa in cfg["EMPRESAS"].keys():
        # usa a chave como nome exibido; valor e caminho eh usado apenas no backend
        empresas[nome_empresa] = nome_empresa
    return empresas


def carregar_mes_ano_default():
    cfg = configparser.ConfigParser()
    cfg.optionxform = str
    if INI_PATH.exists():
        cfg.read(INI_PATH, encoding="utf-8")
    return cfg.get("GERAL", "MES_ANO", fallback="11-2025")


class StatusWindow:
    def __init__(self, root, on_rpa, titulo="Conciliacao"):
        self.root = root
        self.on_rpa = on_rpa

        self.empresas = carregar_empresas()
        displays = list(self.empresas.keys())
        self.selected_empresa = tk.StringVar(value=displays[0])
        self.mes_ano_var = tk.StringVar(value=carregar_mes_ano_default())

        self.root.title(titulo)
        self.root.geometry("")
        self.root.minsize(400, 250)

        self.main_label = ttk.Label(root, text="Aguardando inicio...", font=("Segoe UI", 11, "bold"))
        self.main_label.pack(pady=(10, 5))

        ttk.Label(root, text="Selecionar Empresa:").pack(pady=(5, 0))
        self.empresa_selector = ttk.Combobox(
            root, values=displays, textvariable=self.selected_empresa, state="readonly", width=30, justify="center"
        )
        self.empresa_selector.pack(pady=(0, 8), anchor="center")

        ttk.Label(root, text="Selecione a pasta (MM/AAAA):").pack(pady=(0, 0))
        self.mes_ano_entry = ttk.Entry(root, textvariable=self.mes_ano_var, width=12, justify="center")
        self.mes_ano_entry.pack(pady=(0, 10))

        ttk.Label(root, text="Progresso").pack(pady=(5, 0))
        self.overall_progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        self.overall_progress.pack(padx=20)

        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(pady=(10, 10))

        self.start_rpa_button = ttk.Button(
            self.button_frame, text="Gerar Conciliacao", command=self.start_rpa, width=20
        )
        self.start_rpa_button.pack(side="left", padx=10)

        self.close_button = ttk.Button(
            self.button_frame, text="Fechar", command=self.root.destroy, state="normal", width=20
        )
        self.close_button.pack(side="right", padx=10)

    def _ui(self, func, *args, **kwargs):
        self.root.after(0, func, *args, **kwargs)

    def _lock_buttons(self):
        self._ui(lambda: self.start_rpa_button.config(state="disabled"))
        self._ui(lambda: self.close_button.config(state="disabled"))

    def _unlock_buttons(self):
        self._ui(lambda: self.start_rpa_button.config(state="normal"))
        self._ui(lambda: self.close_button.config(state="normal"))

    def start_rpa(self):
        self._lock_buttons()
        display = self.empresa_selector.get()
        codigo = self.empresas[display]
        mes_ano = self.get_mes_ano()
        self.update_main_label(f"Iniciando {display} ({mes_ano})")
        threading.Thread(target=self._wrap(self.on_rpa, codigo, display, mes_ano), daemon=True).start()

    def _wrap(self, target, codigo, display, mes_ano):
        def runner():
            try:
                target(codigo, display, mes_ano)
            except Exception as exc:
                self.show_popup(f"ERRO: {exc}")
            finally:
                self._unlock_buttons()
        return runner

    def update_main_label(self, message):
        self._ui(lambda: self.main_label.config(text=message))

    def update_progress(self, bar, value):
        self._ui(lambda: bar.config(value=value))

    def finalize(self):
        self.update_main_label("Processo finalizado!")
        self._unlock_buttons()

    def get_mes_ano(self):
        return self.mes_ano_var.get().strip()

    def show_popup(self, message, title="Aviso"):
        self._ui(messagebox.showinfo, title, message)


def criar_janela(on_rpa, titulo="Conciliacao Dominio x Empresa"):
    root = tk.Tk()
    app = StatusWindow(root, on_rpa, titulo=titulo)
    return root, app


if __name__ == "__main__":
    def dummy_rpa(codigo, display, mes_ano):  # pragma: no cover
        app.update_progress(app.overall_progress, 50)
        app.show_popup(f"Rodou {display} ({mes_ano})")
        app.finalize()

    root, app = criar_janela(dummy_rpa, titulo="Demo")
    root.mainloop()
