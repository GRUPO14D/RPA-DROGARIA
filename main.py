from front_base import criar_janela
from conciliacao import run_conciliacao


def rodar_rpa(codigo, display, mes_ano):
    empresa = display
    app.update_main_label(f"Conciliacao em andamento para {empresa} ({mes_ano})")
    app.update_progress(app.overall_progress, 5)
    run_conciliacao(mes_ano, [empresa])
    app.update_progress(app.overall_progress, 100)
    app.update_main_label("Processo finalizado.")


if __name__ == "__main__":
    root, app = criar_janela(rodar_rpa, titulo="Conciliacao Dominio x Empresa")
    app.start_rpa_button.config(text="Gerar Conciliacao")
    root.mainloop()
