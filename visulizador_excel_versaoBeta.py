from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserIconView
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
import pandas as pd

class VisualizadorKivyExcelApp(App):
    def build(self):
        self.title = "Visualizador de Excel com Grid Ajustável"  # Define o título da janela

        layout_principal = BoxLayout(orientation='vertical')  # Cria o layout principal vertical

        botao_carregar = Button(text="Carregar Arquivo Excel", size_hint=(1, 0.1))  # Cria botão para carregar arquivo
        botao_carregar.bind(on_press=self.abrir_seletor_arquivos)  # Associa evento de clique ao método
        layout_principal.add_widget(botao_carregar)  # Adiciona o botão ao layout principal

        self.area_rolagem = ScrollView(size_hint=(1, 0.9), do_scroll_x=True, do_scroll_y=True)  # Cria área de rolagem
        self.grid_dados = GridLayout(cols=1, size_hint_y=None, size_hint_x=None)  # Cria grid layout dinâmico
        self.grid_dados.bind(minimum_height=self.grid_dados.setter('height'))  # Ajusta altura dinâmica
        self.grid_dados.bind(minimum_width=self.grid_dados.setter('width'))  # Ajusta largura dinâmica
        self.area_rolagem.add_widget(self.grid_dados)  # Adiciona grid layout à área de rolagem

        layout_principal.add_widget(self.area_rolagem)  # Adiciona área de rolagem ao layout principal

        return layout_principal  # Retorna o layout principal

    def abrir_seletor_arquivos(self, instancia):
        seletor_arquivos = FileChooserIconView(filters=['*.xls', '*.xlsx'])  # Cria seletor de arquivos filtrando Excel

        botao_confirmar = Button(text="Carregar", size_hint=(1, 0.1))  # Cria botão de confirmar seleção
        botao_confirmar.bind(on_press=lambda x: self.carregar_excel(seletor_arquivos.selection))  # Associa clique ao método

        layout_seletor = BoxLayout(orientation='vertical')  # Cria layout vertical para seletor e botão
        layout_seletor.add_widget(seletor_arquivos)  # Adiciona seletor ao layout
        layout_seletor.add_widget(botao_confirmar)  # Adiciona botão ao layout

        self.popup = Popup(title="Selecionar um Arquivo Excel", content=layout_seletor, size_hint=(0.9, 0.9))  # Cria popup
        self.popup.open()  # Abre o popup

    def carregar_excel(self, selecao):
        if not selecao:  # Verifica se um arquivo foi selecionado
            return
        caminho_arquivo = selecao[0]  # Obtém o caminho do arquivo

        df = pd.read_excel(caminho_arquivo)  # Lê o arquivo Excel usando pandas

        self.grid_dados.clear_widgets()  # Limpa widgets existentes no grid
        self.grid_dados.cols = len(df.columns)  # Configura o número de colunas do grid

        largura_colunas = [self.calcular_largura_coluna(col, df[col]) for col in df.columns]  # Calcula largura das colunas

        for i, coluna in enumerate(df.columns):
            largura_coluna = int(largura_colunas[i])  # Converte largura para int
            self.grid_dados.add_widget(Label(text=str(coluna), size_hint=(None, None), size=(largura_coluna, 40), width=largura_coluna))  # Adiciona cabeçalho

        for linha in df.itertuples(index=False):
            for i, celula in enumerate(linha):
                largura_celula = int(largura_colunas[i])  # Converte largura para int
                self.grid_dados.add_widget(Label(text=str(celula), size_hint=(None, None), size=(largura_celula, 30), width=largura_celula))  # Adiciona células de dados

        self.popup.dismiss()  # Fecha o popup

    def calcular_largura_coluna(self, nome_coluna, coluna):
        maximo_caracteres = max(coluna.apply(lambda x: len(str(x))).max(), len(nome_coluna))  # Calcula largura máxima
        return maximo_caracteres * 10  # Multiplica por 10 para definir a largura

if __name__ == '__main__':
    VisualizadorKivyExcelApp().run()  # Executa a aplicação
