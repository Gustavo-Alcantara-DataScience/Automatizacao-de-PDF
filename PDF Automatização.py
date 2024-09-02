import fitz
import os
import pandas as pd
import matplotlib.pyplot as plt
from kivy.app import App
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.popup import Popup
from kivy.uix.filechooser import FileChooserIconView


class DirectoryChooserPopup(Popup):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.selected_file = None  # Variável para armazenar o arquivo selecionado
        current_path = os.getcwd()
        self.filechooser = FileChooserIconView(path=current_path, dirselect=False)
        self.filechooser.bind(on_submit=self.on_selection)
        self.add_widget(self.filechooser)

    def on_selection(self, filechooser, selection, touch):
        if selection:
            selected_directory = selection[0]
            print(f'Diretório selecionado: {selected_directory}')
            self.dismiss()
            df = pd.read_excel(selected_directory, header=None)  # Leitura do arquivo

            nome = df.iloc[0]
            nome = nome[0]
            self.nome = nome.split("Relatório mensal")[0]

            data = df.iloc[3]
            data = data[0]

            meses = {
                "01": "JANEIRO",
                "02": "FEVEREIRO",
                "03": "MARÇO",
                "04": "ABRIL",
                "05": "MAIO",
                "06": "JUNHO",
                "07": "JULHO",
                "08": "AGOSTO",
                "09": "SETEMBRO",
                "10": "OUTUBRO",
                "11": "NOVEMBRO",
                "12": "DEZEMBRO"
            }

            mes_num = data.split("-")[1]
            self.mes_nome = meses[mes_num]

            self.dados_de_geracao = df.iloc[17].tolist()
            self.dados_de_geracao = self.dados_de_geracao[3:]  # geração diária
            self.geracao_total = self.dados_de_geracao.pop()  # geração total do mês
            self.geracao_total = float(self.geracao_total)
            self.economia = round(0.9 * self.geracao_total, 2)

            self.energia_total = df.iloc[8].tolist()  # geração de todos os meses até então
            self.energia_total = self.energia_total[5]
            self.economia_total = float(self.energia_total)
            self.economia_total = round(0.9 * self.energia_total, 2)
            print(f'Dados da linha 18: {self.dados_de_geracao} e geração total: {self.geracao_total}')
            print(f'Dados da linha 9: {self.energia_total}')

            self.co2 = df.iloc[11].tolist()  # co2
            self.co2 = self.co2[5]
            self.co2 = float(self.co2)

            self.arvores = round(self.co2 / 22)  # arvores

            self.fig, self.ax = plt.subplots()
            self.barras = plt.barh(range(1, len(self.dados_de_geracao) + 1), self.dados_de_geracao, color=(1, 0.5, 0))
            plt.rcParams['font.serif'] = ['Times New Roman']
            plt.xlabel('Energia Gerada(kWh)')
            plt.ylabel('Dia')
            plt.xticks(range(1, len(self.dados_de_geracao), 5))

            for barra in self.barras:
                largura = barra.get_width()
                self.ax.text(largura, barra.get_y() + barra.get_height() / 2, f'{largura}', va='center', ha='left', fontsize='7')

            plt.savefig('geracao_grafico.png', dpi=300, transparent=True)

            self.pdf = fitz.open()
            self.pagina = self.pdf.new_page(width=595, height=842)
            self.pagina.insert_image(self.pagina.rect, filename="C:/Users/Gustavo Alcântara/Desktop/Codigos/Automatização PDF/ModeloSlide_v2.png")
            rect = fitz.Rect(3, 190, 480, 710)
            self.pagina.insert_image(rect, filename="C:/Users/Gustavo Alcântara/Desktop/Codigos/Automatização PDF/geracao_grafico.png")

            return selected_directory

class Manager(ScreenManager):
    pass

class Menu(Screen):
    def open_directorychooser(self):
        self.popup = DirectoryChooserPopup()
        self.popup.open()

    def generate_pdf(self):
        desempenho_inserido = self.ids.desempenho_input.text
        inversor_inserido = self.ids.inversor_input.text
        status_inserido = self.ids.status_input.text
        parecer_inserido = self.ids.parecer_input.text
        
        # Continue a partir daqui para inserir o nome no PDF
        if hasattr(self, 'popup'):
            self.popup.pdf = fitz.open()
            self.popup.pagina = self.popup.pdf.new_page(width=595, height=842)
            self.popup.pagina.insert_image(self.popup.pagina.rect, filename="C:/Users/Gustavo Alcântara/Desktop/Codigos/Automatização PDF/ModeloSlide_v2.png")
            rect = fitz.Rect(3, 190, 480, 710)
            self.popup.pagina.insert_image(rect, filename="C:/Users/Gustavo Alcântara/Desktop/Codigos/Automatização PDF/geracao_grafico.png")

            # Aqui você insere o nome capturado no PDF
            self.popup.pagina.insert_text((389.5, 80.5), inversor_inserido, fontsize=9, fontname='helv', color=(0, 0, 0))
            self.popup.pagina.insert_text((380, 95), status_inserido, fontsize=9, fontname='helv', color=(0, 0, 0))
            self.popup.pagina.insert_text((440, 202), desempenho_inserido, fontsize=20, fontname='helv', color=(0, 0, 0))
            self.popup.pagina.insert_text((440, 202), desempenho_inserido, fontsize=20, fontname='helv', color=(0, 0, 0))
            self.popup.pagina.insert_text((22, 719), parecer_inserido, fontsize=16, fontname='helv', color=(0.5, 0.5, 0.5))

            # Insere os outros dados como antes
            self.popup.pagina.insert_text((380, 66), self.popup.nome, fontsize=9, fontname='helv', color=(0, 0, 0))
            self.popup.pagina.insert_text((78, 202), str(self.popup.geracao_total), fontsize=20, fontname='helv', color=(0, 0, 0))
            self.popup.pagina.insert_text((270, 202), str(self.popup.economia), fontsize=20, fontname='helv', color=(0, 0, 0))
            self.popup.pagina.insert_text((496, 626), str(self.popup.economia_total), fontsize=14, fontname='helv', color=(0.5, 0.5, 0.5))
            self.popup.pagina.insert_text((497, 397), str(self.popup.co2), fontsize=14, fontname='helv', color=(0.5, 0.5, 0.5))
            self.popup.pagina.insert_text((506, 519), str(self.popup.arvores), fontsize=14, fontname='helv', color=(0.5, 0.5, 0.5))
            self.popup.pagina.insert_text((492, 18), str(self.popup.mes_nome), fontsize=14, fontfile='figbo', color=(1, 1, 1))

            self.popup.pdf.save("Relatório Mensal.pdf")
            print("PDF gerado com sucesso!")

            os.startfile("Relatório Mensal.pdf")
            
class Testv(App):
    def build(self):
        return Manager()

Testv().run()
