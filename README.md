Gerador de Relatórios em PDF com Interface Gráfica
Este projeto é um aplicativo desenvolvido em Python usando Kivy, Matplotlib e PyMuPDF (fitz), que permite selecionar um arquivo Excel, extrair dados específicos e gerar um relatório mensal em PDF com gráficos e informações personalizadas.

Funcionalidades
Seleção de Arquivo: O usuário pode selecionar um arquivo Excel através de um popup de seleção de arquivos.
Extração de Dados: O aplicativo extrai dados específicos do arquivo Excel selecionado, como nome, data, dados de geração de energia, economia de CO2, e outros.
Criação de Gráficos: Um gráfico de barras horizontal é gerado com os dados de geração diária de energia.
Geração de PDF: Um PDF é criado com todas as informações extraídas, incluindo gráficos e textos inseridos pelo usuário.
Interface Gráfica: A aplicação possui uma interface gráfica amigável, onde o usuário pode inserir informações adicionais que serão incluídas no relatório PDF.
Requisitos
Python 3.x
Bibliotecas Python:
kivy
fitz (PyMuPDF)
pandas
matplotlib
Um arquivo Excel com a estrutura de dados esperada pelo aplicativo.
Instalação
Clone este repositório:
bash
Copiar código
git clone https://github.com/seu-usuario/gerador-relatorios-pdf.git
Navegue até o diretório do projeto:
bash
Copiar código
cd gerador-relatorios-pdf
Instale as dependências necessárias:
bash
Copiar código
pip install kivy pymupdf pandas matplotlib
Como Usar
Execute o aplicativo:
bash
Copiar código
python main.py
Na interface do aplicativo, clique no botão para abrir o seletor de arquivos e escolha o arquivo Excel que deseja usar.
Preencha os campos de texto com as informações desejadas (desempenho, inversor, status, parecer).
Clique em "Gerar PDF" para criar o relatório. O PDF será salvo no diretório atual com o nome "Relatório Mensal.pdf".
Estrutura do Projeto
main.py: Arquivo principal que contém todo o código da aplicação.
ModeloSlide.png: Imagem usada como modelo de fundo para o relatório PDF.
geracao_grafico.png: Gráfico gerado automaticamente a partir dos dados do arquivo Excel.
Personalização
Fontes e Cores: As fontes e cores dos textos inseridos no PDF podem ser personalizadas alterando os parâmetros das funções insert_text.
Gráficos: O estilo e as cores do gráfico podem ser ajustados na parte do código onde o gráfico é gerado utilizando Matplotlib.
Contribuição
Sinta-se à vontade para fazer um fork deste projeto, adicionar melhorias e enviar pull requests.
