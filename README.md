# Visualizador de Excel com Kivy

Este repositório contém um aplicativo desenvolvido com [Kivy](https://kivy.org/) para visualizar arquivos Excel de forma dinâmica e interativa. A aplicação oferece uma interface simples e intuitiva para carregar planilhas (.xls, .xlsx) e exibir os dados em um grid ajustável com largura de colunas adaptada ao conteúdo.

## Funcionalidades

- **Carregamento de Arquivos Excel**: Permite selecionar e carregar arquivos Excel diretamente da interface do aplicativo.
- **Grid de Dados Dinâmico**: Exibição dos dados em formato de tabela, com largura de colunas ajustada ao conteúdo para melhorar a legibilidade.
- **Área de Rolagem**: Suporte para rolagem horizontal e vertical, facilitando a navegação em planilhas grandes.
- **Interface Simples**: Design intuitivo usando Kivy, com suporte para desktops e dispositivos móveis.

## Pré-requisitos

Para executar o projeto, você precisará de:

- Python 3.7 ou superior.
- As bibliotecas abaixo instaladas:
  - [Kivy](https://kivy.org/doc/stable/gettingstarted/installation.html)
  - [Pandas](https://pandas.pydata.org/)

## Como Executar

1. Clone o repositório:

   ```bash
   git clone https://github.com/seu-usuario/seu-repositorio.git
   cd seu-repositorio
   ```

2. Instale as dependências necessárias:

   ```bash
   pip install kivy pandas openpyxl
   ```

3. Execute o aplicativo:

   ```bash
   python main.py
   ```

4. A interface será aberta e você poderá carregar seus arquivos Excel para visualização.

## Estrutura do Código

- **`VisualizadorKivyExcelApp`**: Classe principal do aplicativo, responsável por construir a interface e gerenciar as interações do usuário.
- **`abrir_seletor_arquivos`**: Método que exibe um popup para seleção de arquivos Excel.
- **`carregar_excel`**: Método que lê e exibe os dados da planilha no grid dinâmico.
- **`calcular_largura_coluna`**: Método para calcular a largura ideal das colunas com base no conteúdo.

## Exemplo de Uso

- Ao iniciar o aplicativo, clique no botão **"Carregar Arquivo Excel"**.
- Selecione um arquivo Excel válido (.xls ou .xlsx).
- O conteúdo será exibido em uma tabela, com suporte para rolagem horizontal e vertical.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias ou novas funcionalidades.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).
