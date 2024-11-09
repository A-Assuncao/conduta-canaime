import tkinter as tk
from tkinter import filedialog, messagebox
from queue import Queue, Empty
import threading
import datetime
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from login_canaime import Login
from playwright.sync_api import sync_playwright

# Variáveis de URLs e unidades
lista_ups = ['PAMC', 'CPBV', 'CPFBV', 'CPP', 'UPRRO']
chamada = 'https://canaime.com.br/sgp2rr/areas/impressoes/UND_ChamadaFOTOS_todos2.php?id_und_prisional='
certidao = 'https://canaime.com.br/sgp2rr/areas/impressoes/UND_CertidaoCarceraria.php?id_cad_preso='

# Inicialização do contexto global do Playwright e do navegador
# Esses objetos serão utilizados durante todo o script
playwright = sync_playwright().start()
browser = playwright.chromium.launch(headless=False)


# Função principal para execução da coleta de dados no Playwright em uma thread separada
def execute_playwright_task(selected_units, queue, stop_event):
    """
    Executa a coleta de dados de cada unidade selecionada. Navega até a página de cada unidade prisional, coleta a
    lista de presos e, em seguida, acessa a página de cada preso para obter informações detalhadas, incluindo a conduta.

    Parâmetros:
    - selected_units (list): Lista das unidades prisionais selecionadas pelo usuário.
    - queue (Queue): Fila para enviar mensagens e atualizações de progresso para a interface gráfica.
    - stop_event (Event): Evento para sinalizar a interrupção da execução se o usuário fechar a janela.
    """
    all_units_data = {}  # Dicionário para armazenar os dados de todas as unidades processadas
    try:
        # Realiza login e abre a página inicial do sistema Canaimé
        page = Login()  # Usando Login() para retornar `page` do Playwright

        def coletar_conduta(cdg, nome):
            """
            Navega até a página de certidão do preso e coleta sua conduta.

            Parâmetros:
            - cdg (str): Código do preso.
            - nome (str): Nome do preso.

            Retorna:
            - str: Conduta do preso coletada na página de certidão.
            """
            page.goto(certidao + cdg)
            page.wait_for_load_state('networkidle')
            conduta = page.locator('tr:nth-child(11) .titulo12bk+ .titulobk').text_content()
            return conduta

        # Itera sobre cada unidade selecionada e coleta a lista de presos e suas condutas
        for unit in selected_units:
            # Interrompe o loop se o evento de parada for acionado
            if stop_event.is_set():
                break

            # Envia mensagem de progresso para a interface
            queue.put(f"Navegando para a unidade {unit}...")
            page.goto(chamada + unit, timeout=60000)
            page.wait_for_load_state('networkidle')

            # Localizadores dos elementos contendo os dados de cada preso na unidade
            tudo = page.locator('.titulobkSingCAPS')
            nomes = page.locator('.titulobkSingCAPS .titulo12bk')
            count = tudo.count()  # Conta o número de presos listados
            unit_list = []

            # Coleta as informações básicas de cada preso na unidade
            for i in range(count):
                if stop_event.is_set():  # Interrompe se o evento de parada for acionado
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

            # Coleta a conduta de cada preso acessando a página de certidão individualmente
            unit_data = []
            for idx, (cdg, ala, nome) in enumerate(unit_list):
                if stop_event.is_set():  # Interrompe se o evento de parada for acionado
                    break

                try:
                    conduta = coletar_conduta(cdg, nome)
                    restantes = total_presos - (idx + 1)
                    queue.put(f"{cdg} - {nome}, Conduta {conduta}, restam {restantes} presos")
                    unit_data.append((cdg, ala, nome, conduta))
                except Exception as e:
                    queue.put(f"Erro ao coletar conduta do preso {nome} (Código: {cdg}): {str(e)}")
                    continue

            # Armazena os dados coletados para a unidade atual
            all_units_data[unit] = unit_data

        # Envia os dados completos para a interface gráfica
        queue.put(all_units_data)

    except Exception as e:
        queue.put(f"Erro no Playwright: {str(e)}")

    finally:
        stop_event.set()


# Função para salvar os dados coletados em um arquivo Excel
def salvar_excel(all_units_data):
    """
    Salva os dados de todas as unidades em um arquivo Excel, criando uma aba para cada unidade.

    Parâmetros:
    - all_units_data (dict): Dicionário com os dados de todas as unidades processadas.
    """
    wb = Workbook()
    wb.remove(wb.active)  # Remove a aba padrão

    # Adiciona uma aba para cada unidade, preenchendo com os dados de cada preso
    for up, data in all_units_data.items():
        df = pd.DataFrame(data, columns=["Código", "Ala", "Preso", "Conduta"])
        ws = wb.create_sheet(title=up)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

    # Define o nome sugerido para o arquivo com a data atual
    data_atual = datetime.datetime.now().strftime("%d-%m-%Y")
    sugestao_nome = f"Lista Conduta {data_atual}.xlsx"
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    caminho_arquivo = filedialog.asksaveasfilename(parent=root, initialdir="~/Desktop", title="Salvar Arquivo",
                                                   defaultextension=".xlsx", initialfile=sugestao_nome,
                                                   filetypes=[("Arquivo Excel", "*.xlsx")])
    root.destroy()

    # Salva o arquivo ou mostra aviso se o usuário cancelar
    if caminho_arquivo:
        wb.save(caminho_arquivo)
        messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")
    else:
        messagebox.showwarning("Ação Cancelada", "O arquivo não foi salvo.")


# Função que verifica a fila de mensagens e atualiza a interface com o progresso
def verificar_fila(root, label_loading):
    """
    Verifica periodicamente a fila para atualizar a interface gráfica com mensagens de progresso
    ou resultados de erro.

    Parâmetros:
    - root (Tk): Objeto raiz da interface gráfica.
    - label_loading (Label): Label da interface para exibir mensagens de progresso.
    """
    if stop_event.is_set():  # Interrompe se o evento de parada for acionado
        return

    try:
        # Verifica se há uma nova mensagem na fila
        message = queue.get_nowait()
        if isinstance(message, dict):
            salvar_excel(message)
            label_loading.config(text="Processo concluído!")
        elif isinstance(message, str):
            label_loading.config(text=message)
        root.update_idletasks()
    except Empty:
        pass
    root.after(100, lambda: verificar_fila(root, label_loading))


# Função que configura e exibe a interface gráfica de seleção de unidades prisionais
def selecionar_unidades():
    """
    Cria e exibe a interface gráfica para seleção de unidades prisionais.
    Após o clique no botão de confirmação, desativa as seleções e inicia o processamento Playwright.
    """
    root = tk.Tk()
    root.title("Seleção de Unidades Prisionais")
    root.geometry("400x300")
    root.eval('tk::PlaceWindow . center')  # Centraliza a janela
    root.attributes('-topmost', True)

    # Variáveis para armazenar seleção das unidades
    unidades_vars = {up: tk.BooleanVar() for up in lista_ups}

    # Adiciona checkboxes para cada unidade com margem
    tk.Label(root, text="Selecione as unidades prisionais:", font=("Arial", 12)).pack(pady=10)
    frame_units = tk.Frame(root)
    frame_units.pack(anchor='w', padx=20)
    checkbuttons = []
    for up, var in unidades_vars.items():
        cb = tk.Checkbutton(frame_units, text=up, variable=var)
        cb.pack(anchor='w')
        checkbuttons.append(cb)

    # Label de carregamento para exibir o progresso com fonte reduzida
    label_loading = tk.Label(root, text="", font=("Arial", 8))
    label_loading.pack(pady=10)

    # Função de confirmação da seleção
    def confirmar_selecao():
        unidades_selecionadas = [up for up, var in unidades_vars.items() if var.get()]
        if not unidades_selecionadas:
            messagebox.showwarning("Nenhuma Unidade Selecionada", "Por favor, selecione pelo menos uma unidade.")
        else:
            for cb in checkbuttons:
                cb.config(state="disabled")
            btn_confirmar.config(state="disabled")
            threading.Thread(target=execute_playwright_task, args=(unidades_selecionadas, queue, stop_event)).start()
            verificar_fila(root, label_loading)

    # Função de fechamento para encerramento completo
    def fechar_janela():
        if messagebox.askokcancel("Fechar", "Deseja realmente fechar sem salvar?"):
            stop_event.set()  # Define o evento de parada
            root.destroy()  # Fecha a janela Tkinter

    # Botão de confirmação e fechamento
    btn_confirmar = tk.Button(root, text="Confirmar", command=confirmar_selecao)
    btn_confirmar.pack(pady=20)
    root.protocol("WM_DELETE_WINDOW", fechar_janela)

    root.mainloop()


# Inicialização da fila e eventos globais
queue = Queue()
stop_event = threading.Event()

# Executa a interface gráfica
selecionar_unidades()

# Encerramento do Playwright e navegador após o fechamento da janela Tkinter
try:
    browser.close()
    playwright.stop()
except Exception as e:
    print(f"Erro ao fechar o Playwright: {str(e)}")
