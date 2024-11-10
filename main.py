import tkinter as tk
from tkinter import filedialog, messagebox
from queue import Queue, Empty
import threading
import datetime
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from login import Login
from playwright.sync_api import sync_playwright
import sys

# Variáveis de URLs e unidades
lista_ups = ['PAMC', 'CPBV', 'CPFBV', 'CPP', 'UPRRO']
chamada = 'https://canaime.com.br/sgp2rr/areas/impressoes/UND_ChamadaFOTOS_todos2.php?id_und_prisional='
certidao = 'https://canaime.com.br/sgp2rr/areas/impressoes/UND_CertidaoCarceraria.php?id_cad_preso='

playwright = sync_playwright().start()
browser = playwright.chromium.launch(headless=False)


def execute_playwright_task(selected_units, queue, stop_event):
    """
    Executa a coleta de dados de cada unidade selecionada. Navega até a página de cada unidade prisional, coleta a
    lista de presos e, em seguida, acessa a página de cada preso para obter informações detalhadas, incluindo a conduta.
    """
    all_units_data = {}
    try:
        # Chama o Login e verifica se ele retorna None (fechou sem autenticar)
        page = Login(test=False)
        if page is None:  # Se Login retornar None, encerra o script
            sys.exit()  # Encerra o script

        def coletar_conduta(cdg, nome):
            page.goto(certidao + cdg)
            page.wait_for_load_state('networkidle')
            conduta = page.locator('tr:nth-child(11) .titulo12bk+ .titulobk').text_content()
            return conduta

        for unit in selected_units:
            if stop_event.is_set():
                break

            queue.put(f"Navegando para a unidade {unit}...")
            page.goto(chamada + unit, timeout=60000)
            page.wait_for_load_state('networkidle')

            tudo = page.locator('.titulobkSingCAPS')
            nomes = page.locator('.titulobkSingCAPS .titulo12bk')
            count = tudo.count()
            unit_list = []

            for i in range(count):
                if stop_event.is_set():
                    break
                try:
                    tudo_tratado = tudo.nth(i).text_content().replace(" ", "").strip()
                    [codigo, _, _, _, alas] = tudo_tratado.split('\n')
                    preso = nomes.nth(i).text_content().strip()
                    cdg = codigo[2:]
                    ala = alas[-3:]
                    unit_list.append((cdg, ala, preso))
                except Exception as e:
                    queue.put(f"Erro ao coletar lista de presos da unidade {unit}, índice {i}: {str(e)}")
                    continue

            total_presos = len(unit_list)
            queue.put(f"Buscando conduta carcerária de um total de {total_presos} presos")

            unit_data = []
            for idx, (cdg, ala, nome) in enumerate(unit_list):
                if stop_event.is_set():
                    break
                try:
                    conduta = coletar_conduta(cdg, nome)
                    restantes = total_presos - (idx + 1)
                    queue.put(f"{cdg} - {nome}, Conduta {conduta}, restam {restantes} presos")
                    unit_data.append((cdg, ala, nome, conduta))
                except Exception as e:
                    queue.put(f"Erro ao coletar conduta do preso {nome} (Código: {cdg}): {str(e)}")
                    continue

            all_units_data[unit] = unit_data

        queue.put(all_units_data)

    except Exception as e:
        queue.put(f"Erro no Playwright: {str(e)}")

    finally:
        stop_event.set()
        try:
            browser.close()
            playwright.stop()
        except Exception as e:
            queue.put(f"Erro ao fechar o Playwright: {str(e)}")


def salvar_excel(all_units_data):
    """
    Salva os dados de todas as unidades em um arquivo Excel, criando uma aba para cada unidade.
    """
    wb = Workbook()
    wb.remove(wb.active)

    for up, data in all_units_data.items():
        df = pd.DataFrame(data, columns=["Código", "Ala", "Preso", "Conduta"])
        ws = wb.create_sheet(title=up)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

    data_atual = datetime.datetime.now().strftime("%d-%m-%Y")
    sugestao_nome = f"Lista Conduta {data_atual}.xlsx"
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    caminho_arquivo = filedialog.asksaveasfilename(parent=root, initialdir="~/Desktop", title="Salvar Arquivo",
                                                   defaultextension=".xlsx", initialfile=sugestao_nome,
                                                   filetypes=[("Arquivo Excel", "*.xlsx")])
    root.destroy()

    if caminho_arquivo:
        wb.save(caminho_arquivo)
        messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")
    else:
        messagebox.showwarning("Ação Cancelada", "O arquivo não foi salvo.")


def verificar_fila(root, label_loading):
    """
    Verifica periodicamente a fila para atualizar a interface gráfica com mensagens de progresso
    ou resultados de erro.
    """
    if stop_event.is_set():
        return

    try:
        message = queue.get_nowait()
        if message == "login_failed":
            sys.exit()  # Encerra o script se o login falhar
        elif isinstance(message, dict):
            salvar_excel(message)
            label_loading.config(text="Processo concluído!")
        elif isinstance(message, str):
            label_loading.config(text=message)
        root.update_idletasks()
    except Empty:
        pass
    root.after(100, lambda: verificar_fila(root, label_loading))


def selecionar_unidades():
    """
    Cria e exibe a interface gráfica para seleção de unidades prisionais.
    """
    global root, label_loading, checkbuttons, btn_confirmar, unidades_vars

    root = tk.Tk()
    root.title("Seleção de Unidades Prisionais")
    root.geometry("400x300")
    root.eval('tk::PlaceWindow . center')
    root.attributes('-topmost', True)

    unidades_vars = {up: tk.BooleanVar() for up in lista_ups}

    tk.Label(root, text="Selecione as unidades prisionais:", font=("Arial", 12)).pack(pady=10)
    frame_units = tk.Frame(root)
    frame_units.pack(anchor='w', padx=20)
    checkbuttons = []
    for up, var in unidades_vars.items():
        cb = tk.Checkbutton(frame_units, text=up, variable=var)
        cb.pack(anchor='w')
        checkbuttons.append(cb)

    label_loading = tk.Label(root, text="", font=("Arial", 8))
    label_loading.pack(pady=10)

    btn_confirmar = tk.Button(root, text="Confirmar", command=confirmar_selecao)
    btn_confirmar.pack(pady=20)
    root.protocol("WM_DELETE_WINDOW", fechar_janela)

    root.mainloop()


def confirmar_selecao():
    """
    Confirma a seleção de unidades e inicia a tarefa de coleta de dados.
    """
    unidades_selecionadas = [up for up, var in unidades_vars.items() if var.get()]
    if not unidades_selecionadas:
        messagebox.showwarning("Nenhuma Unidade Selecionada", "Por favor, selecione pelo menos uma unidade.")
    else:
        thread = threading.Thread(target=execute_playwright_task, args=(unidades_selecionadas, queue, stop_event))
        thread.start()
        verificar_fila(root, label_loading)


def fechar_janela():
    """
    Função de fechamento para encerramento completo.
    """
    if messagebox.askokcancel("Fechar", "Deseja realmente fechar sem salvar?"):
        stop_event.set()
        root.destroy()
        sys.exit()  # Encerra o script completamente


queue = Queue()
stop_event = threading.Event()

# Executa a interface gráfica
selecionar_unidades()
