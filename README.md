# 📊 Gerador de Relatórios em PDF com Interface Gráfica
Este projeto é um aplicativo desenvolvido em Python usando Kivy, Matplotlib e PyMuPDF (fitz), que permite selecionar um arquivo Excel, extrair dados específicos e gerar um relatório de geração de energia mensal em PDF com gráficos e informações personalizadas, mas pode ser adaptado para outras funções, se necessário.

# 🌟 Funcionalidades
- Seleção de Arquivo: O usuário pode selecionar um arquivo Excel através de um popup de seleção de arquivos.
- Extração de Dados: O aplicativo extrai dados específicos do arquivo Excel selecionado, como nome, data, dados de geração de energia, economia de CO2, e outros.
- Criação de Gráficos: Um gráfico de barras horizontal é gerado com os dados de geração diária de energia.
- Geração de PDF: Um PDF é criado com todas as informações extraídas, incluindo gráficos e textos inseridos pelo usuário.
- Interface Gráfica: A aplicação possui uma interface gráfica amigável, onde o usuário pode inserir informações adicionais que serão incluídas no relatório PDF.
# 🛠️ Requisitos
- Python 3.x
- Bibliotecas Python:
- kivy
- fitz (PyMuPDF)
- pandas
- matplotlib
- Um arquivo Excel com a estrutura de dados esperada pelo aplicativo, mas que facilmente pode ser adaptado para outros casos.
# 🚀 Instalação
Clone este repositório:

- git clone https://github.com/seu-usuario/gerador-relatorios-pdf.git

Navegue até o diretório do projeto:

- cd gerador-relatorios-pdf
  
Instale as dependências necessárias:

- pip install kivy pymupdf pandas matplotlib

# 📋 Como Usar
### Execute o aplicativo:

python PDF Automatização.py

### Selecione o arquivo Excel:

Na interface do aplicativo, clique no botão para abrir o seletor de arquivos e escolha o arquivo Excel que deseja usar.

### Preencha os campos de texto:

Insira as informações desejadas (desempenho, inversor, status, parecer).

### Gere o PDF:

Clique em "Gerar PDF" para criar o relatório. O PDF será salvo no diretório atual com o nome Relatório Mensal.pdf.

# 📂 Estrutura do Projeto

PDF Automatização.py: Arquivo principal que contém todo o código da aplicação.

ModeloSlide_v2.png: Imagem usada como modelo de fundo para o relatório PDF.

geracao_grafico.png: Gráfico gerado automaticamente a partir dos dados do arquivo Excel.

testv.kv: Arquivo Kivy com as propriedades funcionais e estéticas da interface gráfica.

# 🎨 Personalização

Fontes e Cores: As fontes e cores dos textos inseridos no PDF podem ser personalizadas alterando os parâmetros das funções insert_text.

Gráficos: O estilo e as cores do gráfico podem ser ajustados na parte do código onde o gráfico é gerado utilizando Matplotlib.

# 🤝 Contribuição

Sinta-se à vontade para fazer um fork deste projeto, adicionar melhorias e enviar pull requests.
