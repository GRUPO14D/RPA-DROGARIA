from front_base import criar_janela
from conciliacao import run_conciliacao, set_logger


def rodar_rpa(codigo, display, mes_ano):
    empresa = display
    set_logger(app.add_log_message)
    app.update_main_label(f"Conciliação em andamento para {empresa} ({mes_ano})")
    run_conciliacao(mes_ano, [empresa])
    app.update_main_label("Processo finalizado.")


def rodar_auto(codigo, display, mes_ano):
    # Automação não implementada neste fluxo
    app.add_log_message("Automação não disponível para esta conciliação.")


if __name__ == "__main__":
    root, app = criar_janela(rodar_rpa, rodar_auto, titulo="Conciliacao Dominio x Empresa")
    app.start_rpa_button.config(text="Gerar Conciliacao")
    root.mainloop()
