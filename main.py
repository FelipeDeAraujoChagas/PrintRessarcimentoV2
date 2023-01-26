from tkinter import *
from tkinter import filedialog
import pandas as pd
from tkintertable import TableCanvas
from playwright.sync_api import sync_playwright
from tkinter import ttk, messagebox
import openpyxl
import os
#from urllib3.util import timeout


class Telas:

    def __init__(self):

        self.telaApp()
        self.organizarpasta()
        self.telaRotina()

    def organizarpasta(self):

        try:

            pass

        except:

            pass

    def telaApp(self):

        self.master = Tk()
        self.master.title("Analytics & Ressarcimento")
        self.master.geometry("600x402+610+153")
        self.master.iconbitmap(default="icone\\armazem.ico")
        self.master.resizable(FALSE, FALSE)

    def frameTela(self):

        self.fr_tela = Frame(self.master, borderwidth=1, relief='sunken', background='#D9D9D9')
        self.fr_tela.place(x=2, y=2, width=598, height=400)

    def telaLogin(self):

        self.frameTela()

        # Criar caixas de entrada
        self.en_login = Entry(self.fr_tela, bd=2, font=('Calibri', 15), justify=CENTER)
        self.en_login.place(width=299, height=34, x=79, y=144)

        esconda_senha = StringVar()
        self.en_senha = Entry(self.fr_tela, bd=2, textvariable=esconda_senha, show='*', font=('Calibri', 15),
                              justify=CENTER)
        self.en_senha.place(width=299, height=34, x=79, y=244)

        # Label
        self.lb_matricula = Label(self.fr_tela, font=('Calibri', 15), text='Matricula', foreground='#fff', background='#303495')
        self.lb_matricula.place(x=80, y=110)

        self.lb_senha = Label(self.fr_tela,font=('Calibri', 15), text='Senha', foreground='#fff', background='#303495')
        self.lb_senha.place(x=80, y=210)

        # Criar botões
        bt_entrar = Button(self.fr_tela, bd=2, text='Entrar',bg='#FFF', command=self.telaRotina)
        bt_entrar.place(width=96, height=32, x=187, y=319)

        self.aparecer()

    def telaRotina(self):

        self.frameTela()

        img_botao1 = PhotoImage(file="imagem\\bt_upload.png")

        self.fr_tela2 = Frame(self.master, borderwidth=1, relief='sunken', background='#D9D9D9')
        self.fr_tela2.place(x=20, y=120, width=565, height=175)

        bt_upload = Button(self.fr_tela, image=img_botao1, command=self.upLoudArquivo)
        bt_upload.place(x=500, y=70, width=41, height=30)

        bt_execulta = Button(self.fr_tela, bd=2, text='Execulta', bg='#FFF', command=self.execultarRotinaBraser)
        bt_execulta.place(x=275, y=330, width=70, height=32)

        # Label
        self.lb_endereco = Label(self.fr_tela, text='', background='#fff', anchor=E, bd=2)
        self.lb_endereco.place(x=38, y=70, width=460, height=30)

        self.aparecer()

    def aparecer(self):

        self.master.mainloop()

    def upLoudArquivo(self):

        try:

            self.path = filedialog.askopenfilename(filetypes=(('Arq Excel', '*.xls*'), ('All files', '*.*')))
            self.lb_endereco['text'] = self.path
            self.data_fr = pd.read_excel(self.path)
            em = ManipulaArquivo()
            em.criarArqCSV(self.path)
            dire = r'dado.csv'
            table = TableCanvas(self.fr_tela2)
            table.importCSV(dire, sep=';')
            table.show()

        except:

            pass

    def execultarRotinaBraser(self):

        try:

            rodar = PegarPrint()
            rodar.programaPrincipal()

        except Exception as e:

            msg = str(e)
            messagebox.showinfo("informação", msg)


class ManipulaArquivo:

    def criarArqCSV(self, path):

        caminho = path
        list_df = pd.read_excel(caminho)
        dados = list_df[["Pedido", "Instância", "item", "MOTIVO DA COBRANÇA"]]
        df = pd.DataFrame(data=dados)
        df.to_csv('dado.csv', index=False, sep=";", encoding='latin-1')


class PegarPrint:

    def programaPrincipal(self):

        path = "dado.csv"
        site = "http://intranet.suanova.com/transportes/ressarcimento/SitePages/Consultar%20Instancia_v2.aspx"
        self.nome_pasta_guadar = r'\\nas01.via.varejo.corp\gremio01\TRANSPORTE\FINANCEIRO\Clécia Lucena\DEMANDA SEQUOIA\ARQUIVOS'

        #p: Playwright

        with sync_playwright() as p:

            try:
                browser = p.chromium.launch(headless=False, executable_path=r"chrome-win\chrome.exe")

            except Exception as e:

                msg = str(e)
                messagebox.showinfo("informação", msg)


            self.page = browser.new_page()
            self.page.goto(site)
            self.page.set_viewport_size({"width": 1600, "height": 2000})

            self.list_df = pd.read_csv(path, encoding="latin-1", sep=";")
            self.criar_excel()
            self.abrirjanela()

    def rodarloop(self):

        try:

            ttal = int(len(self.list_df["Pedido"]))
            for key, pedido in enumerate(self.list_df["Pedido"]):

                print(pedido)
                porcento = ((key + 1)/ttal) * 100
                self.pd["value"] = key + 1
                self.lb_informa['text'] = f'{key + 1}/{ttal}   {porcento:.2f} %'
                self.app.update()
                self.pedido = pedido
                self.instancia = self.list_df[f"Instância"][key]
                self.item = self.list_df[f"item"][key]
                self.motivo_cobranca = self.list_df[f"MOTIVO DA COBRANÇA"][key]
                if self.page.locator(f"text = CNPJ da Transportadora ").is_visible():
                    self.limpar_pedido_anterior()
                self.preencher_pedido(self.pedido)
                self.pesquisar_pedido()
                self.checar_instancia_pedido()
                self.preencher_excel()

            self.page.close()
            messagebox.showinfo("Fim", "Processo encerrado")

        finally:

            self.salvar_excel()

    def checar_motivoda_cobranca(self):

        motivo_web = self.frame.locator().inner_text().strip()
        if str(motivo_web) != "":
            self.motivo_cobranca = motivo_web

    def nome_pasta_amazenar_arquivo(self):

        self.pasta = str(self.pedido) + " - " + str(self.instancia) + " - " + str(self.item) + " - " + str(self.motivo_cobranca)

    def setar_frame(self):

        self.page.wait_for_timeout(500)
        self.frame = self.page.frame_locator("iframe.ms-dlgFrame")

    def checar_instancia_pedido(self):

        ver_numero_itens = self.page.locator('xpath = //html/body/form/div[5]/div/div[1]/div[5]/div/div/div[4]/div/div/table/tbody/tr/td/div/div/div/div/div/div/div[2]/div/div[1]/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/div/table/tbody/tr[4]/td/table[2]/tbody/tr[2]/td/div/table/tbody/tr')

        for c in range(2, ver_numero_itens.count() + 1):

            linha_tabela_pedido = self.page.locator(f'xpath = //html/body/form/div[5]/div/div[1]/div[5]/div/div/div[4]/div/div/table/tbody/tr/td/div/div/div/div/div/div/div[2]/div/div[1]/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/div/table/tbody/tr[4]/td/table[2]/tbody/tr[2]/td/div/table/tbody/tr[{c}]/td[5]')
            ver_infor = linha_tabela_pedido.inner_text().rstrip()

            if str(ver_infor) == str(self.instancia):

                self.sku = self.checar_sku_instancia(c)

                if str(self.sku) == str(self.item):

                    self.checar_motivo()
                    self.nome_pasta_amazenar_arquivo()
                    self.criar_pastas()
                    self.checar_aquivos()
                    self.page.set_viewport_size({"width": 1900, "height": 2000})
                    self.page.wait_for_timeout(1000)
                    self.tirar_prints(tela="item do pedido ")
                    self.fechar_tela()
                    self.page.set_viewport_size({"width": 1900, "height": 1300})
                    self.page.wait_for_timeout(1000)
                    self.tirar_prints(tela="instancia do pedido ")
                    break

                else:

                    self.fechar_tela()

    def checar_motivo(self):

        self.motivo_cobranca_intranet = self.frame.locator("span#ctl00_ctl21_g_037ceea0_2087_4341_90d6_ea37224bab53_ctl00_lblMOTIVO_COBRANCA").inner_text().rstrip()

        if self.motivo_cobranca_intranet != "":
            self.motivo_cobranca = self.motivo_cobranca_intranet

    def checar_sku_instancia(self, sequencial):

        self.abrir_informacao(sequencial)
        self.expadir_tela()
        self.setar_frame()
        local_item = self.frame.locator("span#ctl00_ctl21_g_037ceea0_2087_4341_90d6_ea37224bab53_ctl00_lblITEM")
        sku = local_item.inner_text().rstrip()

        return sku.strip()

    def checar_aquivos(self):

        try:

            self.retorno_download = "Sem arquivos baixados"
            self.anex_ponto = self.frame.locator("table#ctl00_ctl21_g_037ceea0_2087_4341_90d6_ea37224bab53_ctl00_gdvFiles")
            self.conferir_numero_arquivos()
            self.preencher_excel()

        except:

            self.retorno_download = "Não foram encontrados arquivos para baixar"

    def conferir_numero_arquivos(self):

        n = self.anex_ponto.inner_html(timeout=3000).count('intranet.suanova')

        self.nome_arquivo_anterior = ""
        for c in range(0, n):

            clik = f"xpath = //*[@id='ctl00_ctl21_g_037ceea0_2087_4341_90d6_ea37224bab53_ctl00_gdvFiles']/tbody/tr[{c + 1}]/td[2]/a"
            elemento = self.frame.locator(f"{clik}")
            print(f"arq anterior = {self.nome_arquivo_anterior}", f"arquivo = {elemento.inner_text()}")
            if str(self.nome_arquivo_anterior.strip()) != str(elemento.inner_text().strip()) and str(elemento.inner_text().strip()) != "":
                self.nome_arquivo_anterior = elemento.inner_text().strip()
                self.fazer_download_arquivo(clik)

        if self.retorno_download != "Exitem Download que não foram baixados":

            self.retorno_download = "Todos Download daixados"

    def fazer_download_arquivo(self, clik):

        try:

            with self.page.expect_download() as download_info:
                self.frame.locator(f"{clik}").click()

            download = download_info.value
            nome_arq_baixado = self.tratar_nome_arq(self.frame.locator(f"{clik}").inner_text())
            path = f"{self.nome_pasta_guadar}\\{self.pasta}\\{nome_arq_baixado}"
            download.save_as(path)
            self.page.wait_for_timeout(500)

        except:

            self.retorno_download = "Exitem Download que não foram baixados"

    def tratar_nome_arq(self, text):

        arq = str(text).strip()
        arq = arq.replace("/", "-")
        arq = arq.replace("\\", "-")
        arq = arq.replace(":", "-")
        print(arq)
        return arq

    def tirar_prints(self, tela):

        nome_do_arquivo = f"{self.nome_pasta_guadar}\\" + str(self.pasta) + f"\\{tela}" + str(self.pedido) + ".png"
        self.print_tela(nome_do_arquivo)

    def limpar_pedido_anterior(self):

        bt_cancelar = self.page.locator('input[id = "ctl00_ctl21_g_ba3a473a_1336_4e1e_adcb_5ef3fc254e67_ctl00_ibtCancelarPesquisarInstancias"]')
        bt_cancelar.click()
        self.page.wait_for_timeout(1000)

    def preencher_pedido(self, pedido):

        caixaPesquisa = self.page.locator(
            "input[name='ctl00$ctl21$g_ba3a473a_1336_4e1e_adcb_5ef3fc254e67$ctl00$txtEntrega']")
        caixaPesquisa.click()
        self.page.wait_for_timeout(500)
        caixaPesquisa.fill(f'{pedido}')
        self.page.wait_for_timeout(500)

    def pesquisar_pedido(self):

        botaoPesquisar = self.page.locator('xpath = //*[@id="ctl00_ctl21_g_ba3a473a_1336_4e1e_adcb_5ef3fc254e67_ctl00_ibtConfirmarPesquisarInstancias"]')
        botaoPesquisar.click()
        self.page.locator(f"text = CNPJ da Transportadora ").is_disabled()

    def abrir_informacao(self, sequencial):

        abrirItem = self.page.locator(
           f"xpath = //html/body/form/div[5]/div/div[1]/div[5]/div/div/div[4]/div/div/table/tbody/tr/td/div/div/div/div/div/div/div[2]/div/div[1]/table/tbody/tr/td/table/tbody/tr[2]/td/div/div/div/table/tbody/tr[4]/td/table[2]/tbody/tr[2]/td/div/table/tbody/tr[{sequencial}]/td[1]/a")

        abrirItem.click()
        entao = self.page.frame_locator("iframe.ms-dlgFrame")
        entao.locator("text = Dados de Parcial").is_disabled()

    def expadir_tela(self):

        expandir = self.page.locator('a[title="Maximizar"]')
        expandir.click()
        self.page.wait_for_timeout(500)

    def fechar_tela(self):

        fechar = self.page.locator('a[title="Fechar"]')
        fechar.click()
        self.page.wait_for_timeout(500)

    def print_tela(self, text):

        self.page.screenshot(path=text, full_page=True)
        self.page.wait_for_timeout(500)

    def criar_pastas(self):

        try:

            path = os.path.join(self.nome_pasta_guadar, self.pasta)
            os.makedirs(path)

        except:

            print("pasta não criada")

    def abrirjanela(self):

        self.app = Tk()
        self.app.title("Barra de progresso")
        self.app.geometry("330x200")

        self.pd = ttk.Progressbar(self.app, orient=HORIZONTAL, maximum=len(self.list_df["Pedido"]), mode="determinate")
        self.pd.place(x=30, y=120, width=270, height=30)

        self.lb_informa = Label(self.app, text='', font=("calibri", 15))
        self.lb_informa.place(x=100, y=70)

        self.rodarloop()

        self.app.mainloop()

    def criar_excel(self):

        self.book = openpyxl.Workbook()
        self.sheet = self.book.active
        self.sheet.append(['Pedido', 'Instancia', 'Item', 'motivo da Cobrança', 'Nome da pasta', 'Downloads'])

    def preencher_excel(self):

        self.sheet.append([f'{self.pedido}', f'{self.instancia}', f'{self.item}',
                           f'{self.motivo_cobranca}', f'{self.pasta}', f'{self.retorno_download}'])

    def salvar_excel(self):


        pasta_downloads = r'\\nas01.via.varejo.corp\gremio01\TRANSPORTE\FINANCEIRO\Clécia Lucena\DEMANDA SEQUOIA\RELATORIOS JA FINALIZADOS'
        arquivos = os.listdir(pasta_downloads)
        arq = 'Retorno print ressarcimento.xlsx'
        plan = pasta_downloads + "\\" + arq

        c = 0
        while True:
            c += 1
            sair = True
            for n in arquivos:
                if n == arq:
                    sair = False
                    arq = f'Retorno print ressarcimento({c}).xlsx'
                    plan = pasta_downloads + "\\" + arq

            if sair:
                break

        self.book.save(plan)


if __name__ == "__main__":
    Telas()
