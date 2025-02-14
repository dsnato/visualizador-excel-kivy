from kivymd.app import MDApp
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.button import MDRaisedButton
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivymd.uix.textfield import MDTextField
from kivymd.uix.dialog import MDDialog
from kivy.uix.filechooser import FileChooserIconView
from kivymd.uix.toolbar import MDTopAppBar
import pandas as pd

class VisualizadorKivyMDExcelApp(MDApp):
    def build(self):
        self.title = "Visualizador de Excel com Grid Ajustável"  # Define o título da janela

        self.layout_principal = MDBoxLayout(orientation='vertical', padding=10, spacing=10)  # Adiciona padding e spacing ao layout principal

        toolbar = MDTopAppBar(title="Visualizador de Excel", md_bg_color=(0, 0, 0, 1), elevation=10)  # Preto com sombra
        self.layout_principal.add_widget(toolbar)  # Adiciona a barra de ferramentas ao layout principal

        botao_carregar = MDRaisedButton(text="Carregar Arquivo Excel", size_hint=(1, 0.1), text_color=(1, 1, 1, 1))  # Cria botão para carregar arquivo com texto branco
        botao_carregar.bind(on_press=self.abrir_seletor_arquivos)  # Associa evento de clique ao método
        self.layout_principal.add_widget(botao_carregar)  # Adiciona o botão ao layout principal

        self.area_rolagem = ScrollView(size_hint=(1, 0.9), do_scroll_x=True, do_scroll_y=True, bar_width=20)  # Cria área de rolagem
        self.grid_dados = GridLayout(cols=1, size_hint_y=None, size_hint_x=None, row_default_height=40, row_force_default=True)  # Cria grid layout dinâmico
        self.grid_dados.bind(minimum_height=self.grid_dados.setter('height'))  # Ajusta altura dinâmica
        self.grid_dados.bind(minimum_width=self.grid_dados.setter('width'))  # Ajusta largura dinâmica
        self.area_rolagem.add_widget(self.grid_dados)  # Adiciona grid layout à área de rolagem

        self.layout_principal.add_widget(self.area_rolagem)  # Adiciona área de rolagem ao layout principal

        return self.layout_principal  # Retorna o layout principal

    def abrir_seletor_arquivos(self, instancia):
        seletor_arquivos = FileChooserIconView(filters=['*.xls', '*.xlsx'])  # Cria seletor de arquivos com ícones filtrando Excel

        botao_confirmar = MDRaisedButton(text="Carregar", size_hint=(1, 0.1), text_color=(1, 1, 1, 1))  # Cria botão de confirmar seleção com texto branco
        botao_confirmar.bind(on_press=lambda x: self.carregar_excel(seletor_arquivos.selection))  # Associa clique ao método

        layout_seletor = MDBoxLayout(orientation='vertical', padding=10, spacing=10)  # Cria layout vertical para seletor e botão com padding e spacing
        layout_seletor.add_widget(seletor_arquivos)  # Adiciona seletor ao layout
        layout_seletor.add_widget(botao_confirmar)  # Adiciona botão ao layout

        self.popup = Popup(title="Selecionar um Arquivo Excel", content=layout_seletor, size_hint=(0.9, 0.9))  # Cria popup
        self.popup.open()  # Abre o popup

    def carregar_excel(self, selecao):
        if not selecao:  # Verifica se um arquivo foi selecionado
            return
        caminho_arquivo = selecao[0]  # Obtém o caminho do arquivo

        self.df = pd.read_excel(caminho_arquivo)  # Lê o arquivo Excel usando pandas

        self.remover_campos_filtro()  # Remove os campos de filtro antes de carregar um novo Excel
        self.adicionar_campos_filtro()  # Adiciona os campos de filtro após carregar o Excel

        self.atualizar_grid(self.df)  # Atualiza o grid com os dados do DataFrame

        self.popup.dismiss()  # Fecha o popup

    def remover_campos_filtro(self):
        if hasattr(self, 'input_valor') and self.input_valor:
            self.layout_principal.remove_widget(self.input_valor)
        if hasattr(self, 'botao_filtrar') and self.botao_filtrar:
            self.layout_principal.remove_widget(self.botao_filtrar)

    def adicionar_campos_filtro(self):
        self.input_valor = MDTextField(hint_text="Digite o valor para filtrar", size_hint=(1, 0.1))  # Cria campo de texto para valor de busca
        self.layout_principal.add_widget(self.input_valor)  # Adiciona o campo de texto ao layout principal

        self.botao_filtrar = MDRaisedButton(text="Filtrar", size_hint=(1, 0.1), text_color=(1, 1, 1, 1))  # Cria botão para realizar filtro com texto branco
        self.botao_filtrar.bind(on_press=self.filtrar_por_valor)  # Associa evento de clique ao método de filtro
        self.layout_principal.add_widget(self.botao_filtrar)  # Adiciona o botão ao layout principal

    def atualizar_grid(self, df):
        self.grid_dados.clear_widgets()  # Limpa widgets existentes no grid
        self.grid_dados.cols = len(df.columns)  # Configura o número de colunas do grid

        largura_colunas = [self.calcular_largura_coluna(col, df[col]) for col in df.columns]  # Calcula largura das colunas

        for i, coluna in enumerate(df.columns):
            largura_coluna = int(largura_colunas[i])  # Converte largura para int
            self.grid_dados.add_widget(Label(text=str(coluna), size_hint=(None, None), size=(largura_coluna, 40), width=largura_coluna, color=(0, 0, 0, 1)))  # Adiciona cabeçalho com texto preto

        for linha in df.itertuples(index=False):
            for i, celula in enumerate(linha):
                largura_celula = int(largura_colunas[i])  # Converte largura para int
                self.grid_dados.add_widget(Label(text=str(celula), size_hint=(None, None), size=(largura_celula, 30), width=largura_celula, color=(0, 0, 0, 1)))  # Adiciona células de dados com texto preto

    def calcular_largura_coluna(self, nome_coluna, coluna):
        maximo_caracteres = max(coluna.apply(lambda x: len(str(x))).max(), len(nome_coluna))  # Calcula largura máxima
        return maximo_caracteres * 10  # Multiplica por 10 para definir a largura

    def filtrar_por_valor(self, instancia):
        valor = self.input_valor.text.strip().lower()  # Obtém o valor de busca e remove espaços em branco, convertendo para minúsculas

        try:
            df_filtrado = self.df[self.df.apply(lambda row: row.astype(str).str.contains(valor, case=False).any(), axis=1)]  # Filtra o DataFrame ignorando maiúsculas e minúsculas
            self.atualizar_grid(df_filtrado)  # Atualiza o grid com os dados filtrados
        except Exception as e:
            self.mostrar_erro(f"Erro ao filtrar: {e}")

    def mostrar_erro(self, mensagem):
        self.dialog = MDDialog(title="Erro", text=mensagem, size_hint=(0.8, 0.4), buttons=[MDRaisedButton(text="Fechar", on_release=self.fechar_dialog)])
        self.dialog.open()

    def fechar_dialog(self, instancia):
        self.dialog.dismiss()

if __name__ == '__main__':
    VisualizadorKivyMDExcelApp().run()  # Executa a aplicação
