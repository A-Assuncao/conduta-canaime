
# Conduta Canaimé

Este projeto coleta dados sobre unidades prisionais e gera relatórios de conduta de internos em Excel usando Playwright e Tkinter para automação e interface gráfica.

## Visão Geral

Este script permite a coleta de informações sobre presos em várias unidades prisionais utilizando o sistema Canaimé, e exporta os dados em um arquivo Excel. Ele é configurado para acessar URLs específicas, fazer login no sistema Canaimé, e extrair informações relevantes, como código, ala, nome e conduta dos presos.

## Funcionalidades

-   **Interface Gráfica**: Permite a seleção de unidades prisionais via uma interface com Tkinter.
-   **Automação de Coleta de Dados**: Utiliza o Playwright para acessar páginas específicas do sistema Canaimé e extrair dados.
-   **Exportação para Excel**: Os dados extraídos são salvos em um arquivo Excel, com abas separadas para cada unidade prisional selecionada.
-   **Mensagens de Erro e Sucesso**: Exibe mensagens de aviso e sucesso ao usuário através de `messagebox` do Tkinter.

## Dependências

Certifique-se de que as seguintes bibliotecas estejam instaladas:

-   **Playwright**: Utilizado para automação do navegador.
-   **Pandas**: Para manipulação e criação de dados tabulares.
-   **Openpyxl**: Para criação e edição de arquivos Excel.
-   **Tkinter**: Interface gráfica para seleção de unidades e interação com o usuário.

Para instalar as dependências, você pode usar o seguinte comando:

```bash
pip install pandas openpyxl playwright
```

### Instalando o Playwright

Após instalar a biblioteca Playwright, execute o comando abaixo para instalar o navegador Chromium:


```bash
playwright install
```

## Uso

1.  **Selecione as Unidades**: Ao executar o script, a interface gráfica permitirá a seleção das unidades prisionais desejadas.
2.  **Processo de Automação**: Após confirmar a seleção, o script fará login e coletará dados dos presos.
3.  **Exportação para Excel**: Os dados são exportados para um arquivo Excel, com cada unidade em uma aba separada.

### Execução

Para iniciar o script, execute o arquivo principal com:
```
python main.py
```

## Estrutura do Código

-   `execute_playwright_task`: Função que realiza a coleta de dados usando o Playwright.
-   `salvar_excel`: Função que salva os dados coletados em um arquivo Excel.
-   `selecionar_unidades`: Função principal que inicia a interface gráfica de seleção de unidades e inicia o processo de automação e exportação.

## Personalização

-   **URLs Base**: As URLs para consulta no Canaimé podem ser alteradas nas variáveis `chamada` e `certidao`.
-   **Lista de Unidades**: A lista de unidades prisionais, `lista_ups`, pode ser ajustada para adicionar ou remover unidades.

## Considerações

-   **Timeouts do Playwright**: Se ocorrerem problemas de carregamento, ajuste o valor do parâmetro `timeout` na chamada `page.goto()` no `execute_playwright_task`.
-   **Thread Safety**: As mensagens `messagebox` devem ser exibidas na thread principal do Tkinter para evitar erros de execução.

## Contribuições

Contribuições são bem-vindas! Por favor, envie um pull request com sugestões de melhorias ou correções.
