# SistemaCotacaodeMoedas
 Sistema que permite ter acesso as cotações de moedas
---

# Cotação de Moedas - Aplicação Python com Tkinter

## Objetivo

Este projeto tem como objetivo fornecer uma ferramenta gráfica para obter cotações de moedas, tanto de uma única moeda específica em um determinado dia, quanto para múltiplas moedas ao longo de um período de datas. A aplicação utiliza a API da AwesomeAPI para buscar os valores das moedas em tempo real e permite ao usuário selecionar um arquivo Excel contendo as moedas para atualizar suas cotações em massa.

## Funcionalidades

### 1. **Consulta de Cotação de Uma Única Moeda:**
   - O usuário pode selecionar uma moeda a partir de um menu suspenso (combobox) com a lista completa de moedas disponíveis.
   - É possível escolher uma data específica utilizando um calendário integrado para verificar a cotação daquela moeda no dia selecionado.
   - Ao clicar no botão "Pegar Cotação", o valor da moeda selecionada é exibido na interface, considerando o valor da moeda em relação ao real brasileiro (BRL).
   
### 2. **Consulta de Cotações para Múltiplas Moedas:**
   - O usuário pode importar um arquivo Excel que contenha uma lista de moedas (na primeira coluna).
   - Após selecionar o arquivo, o usuário define um intervalo de datas (data inicial e data final).
   - A aplicação então acessa a API para obter as cotações diárias de todas as moedas listadas no arquivo Excel dentro do intervalo de datas fornecido.
   - As cotações são automaticamente atualizadas na planilha e o arquivo é salvo com as novas cotações.

### 3. **Interface Gráfica Intuitiva:**
   - A aplicação foi desenvolvida com a biblioteca `Tkinter`, proporcionando uma interface gráfica amigável e fácil de usar.
   - Permite que o usuário interaja de maneira simples para realizar consultas tanto para uma única moeda quanto para múltiplas moedas.

### 4. **Exibição de Resultados e Mensagens:**
   - Mensagens de erro são exibidas de maneira clara caso ocorra algum problema, como a falta de um arquivo Excel válido ou falha na obtenção da cotação da moeda.
   - Quando a consulta é bem-sucedida, o valor da cotação ou a confirmação de atualização de cotações é exibida ao usuário.

### 5. **Funcionalidade de Seleção de Arquivo Excel:**
   - O usuário pode selecionar um arquivo Excel com moedas listadas, que será lido e processado pela aplicação.
   - As cotações de todas as moedas no arquivo são atualizadas de acordo com o intervalo de datas fornecido.

### 6. **Salvamento Automático do Arquivo Atualizado:**
   - Após a atualização das cotações, o arquivo Excel é salvo com o nome `Resultado_Cotacoes_Moedas.xlsx`, mantendo os valores das cotações diárias atualizados para cada moeda.

## Tecnologias Utilizadas

- **Python**: Linguagem de programação utilizada para o desenvolvimento da aplicação.
- **Tkinter**: Biblioteca gráfica para criar a interface de usuário.
- **Pandas**: Utilizada para manipulação e atualização do arquivo Excel com cotações de moedas.
- **Requests**: Biblioteca para fazer as requisições HTTP à API da AwesomeAPI.
- **Tkcalendar**: Widget de calendário para facilitar a seleção de datas.
- **AwesomeAPI**: API pública que fornece cotações de moedas em tempo real.

## Como Usar

1. **Instalação das Dependências**:
   - Para executar a aplicação, instale as dependências necessárias com o seguinte comando:

     ```bash
     pip install pandas requests tk tkcalendar
     ```

2. **Rodando a Aplicação**:
   - Após instalar as dependências, execute o script Python diretamente.

     ```bash
     python cotacao_moedas.py
     ```

3. **Interação com a Aplicação**:
   - A interface gráfica permitirá que você selecione uma moeda e uma data para consultar a cotação.
   - Para atualizar múltiplas cotações, basta selecionar um arquivo Excel com as moedas e definir o intervalo de datas.

## Exemplo de Uso

- **Consulta de uma única moeda**: Selecione a moeda desejada e a data, e o valor será exibido na tela.
- **Atualização de múltiplas cotações**: Importe um arquivo Excel com moedas, defina o período e clique para atualizar as cotações. O arquivo será salvo com as cotações atualizadas.

## Conclusão

Esta ferramenta é útil para profissionais, investidores ou qualquer pessoa que precise acompanhar as cotações de diversas moedas de maneira rápida e prática, com a conveniência de uma interface gráfica e a possibilidade de atualizar as cotações em massa a partir de um arquivo Excel.
