# 📊 Visualizador de Excel com KivyMD

Este repositório contém 3 versões de um aplicativo desenvolvido com **Kivy** (versões beta) e **KivyMD** (versão final) para visualizar arquivos Excel de forma dinâmica e interativa. A aplicação oferece uma interface simples e intuitiva para carregar planilhas (`.xls`, `.xlsx`) e exibir e filtrar os dados em um **grid ajustável** com largura de colunas adaptada ao conteúdo.

---

## ✨ Funcionalidades

- 📂 **Carregamento de Arquivos Excel**: Permite selecionar e carregar arquivos Excel diretamente da interface do aplicativo.
- 📊 **Grid de Dados Dinâmico**: Exibição dos dados em formato de tabela, com largura de colunas ajustada ao conteúdo para melhor legibilidade.
- 🔄 **Área de Rolagem**: Suporte para rolagem horizontal e vertical, facilitando a navegação em planilhas grandes.
- 🔍 **Filtro de Dados**: Campo de texto para inserir valores e filtrar os dados da tabela com base no valor especificado.
- 🎨 **Interface Moderna com KivyMD**: Design intuitivo e responsivo, compatível com desktops e dispositivos móveis.
- 📁 **Seletor de Arquivos com Ícones**: Seleção de arquivos Excel utilizando ícones para melhor visualização e desempenho.

---

## 📌 Pré-requisitos

Para executar o projeto, você precisará de:

- **Python 3.7 ou superior**
- As seguintes bibliotecas instaladas:
  
  ```sh
  pip install kivy kivymd pandas openpyxl
  ```

---

## 🚀 Como Executar

1. Clone o repositório:
   ```sh
   git clone https://github.com/dsnato/visualizador-excel-kivy.git
   cd visualizador-excel-kivy
   ```
2. Instale as dependências necessárias:
   ```sh
   pip install kivy kivymd pandas openpyxl
   ```
3. Execute o aplicativo:
   ```sh
   python "visualizador_excel_escolha-a-versão".py
   ```
4. A interface será aberta e você poderá carregar seus arquivos Excel para visualização.

---

## 📂 Estrutura do Código

- **`VisualizadorKivyExcelApp`**: Classe principal do aplicativo, responsável por construir a interface e gerenciar as interações do usuário.
- **`abrir_seletor_arquivos`**: Método que exibe um popup para seleção de arquivos Excel.
- **`carregar_excel`**: Método que lê e exibe os dados da planilha no grid dinâmico.
- **`remover_campos_filtro`**: Método para remover os campos de filtro antes de carregar um novo arquivo Excel.
- **`adicionar_campos_filtro`**: Método para adicionar os campos de filtro após carregar um arquivo Excel.
- **`atualizar_grid`**: Método para atualizar o grid com os dados do DataFrame.
- **`calcular_largura_coluna`**: Método para calcular a largura ideal das colunas com base no conteúdo.
- **`filtrar_por_valor`**: Método para filtrar os dados da tabela com base no valor inserido no campo de filtro.

---

## 📖 Exemplo de Uso

1. **Inicie o aplicativo** e clique no botão **"Carregar Arquivo Excel"**.
2. **Selecione um arquivo Excel válido** (`.xls` ou `.xlsx`).
3. O conteúdo será exibido em uma tabela, com suporte para **rolagem horizontal e vertical**.
4. **Insira um valor no campo de filtro** e clique em **"Filtrar"** para exibir apenas os dados correspondentes.

---

## 🤝 Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para **abrir issues** ou **enviar pull requests** com melhorias ou novas funcionalidades.

---

## 📜 Licença

Este projeto está licenciado sob a **MIT License**.

