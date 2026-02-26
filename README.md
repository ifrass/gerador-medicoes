# Extrator de Planilhas de Análise - Engenharia Civil

Bem-vindo ao **Extrator de Planilhas de Análise**. Esta ferramenta foi desenvolvida para automatizar o processo de extração de quantitativos de projetos de engenharia enviados por prefeituras para licitação.

## O que a ferramenta faz?
Ao receber a planilha de análise padronizada (com abas de CSV, Ruas, Trechos, etc.), o sistema lê o código original e gera um **novo arquivo Excel 100% limpo**.

* **Isolamento de Etapas:** Mantém itens repetidos (ex: Transporte) rigorosamente separados em suas respectivas etapas (Terraplenagem, Drenagem, Urbanização), preservando a lógica orçamentária do projetista.
* **Índice Automático:** Gera uma aba `Lista_de_Ruas` com o índice exato de todos os trechos processados, mantendo a ordem original do arquivo da prefeitura.
* **Cópia Segura:** Prepara os dados matematicamente alinhados (a partir da linha 10, coluna S) para que você possa fazer o "Copiar e Colar Valores" para a Planilha Modelo Oficial do governo sem disparar travamentos do Excel.

## ⚙️ Instalação (Para a Equipe)
Para rodar esta ferramenta no seu computador, você precisará de dois requisitos básicos:
1. **Python** (Pode ser baixado diretamente pela Microsoft Store do Windows).
2. **Git** (Para receber as atualizações automáticas do sistema).

**Primeiro Acesso:**
Abra o terminal/Prompt de Comando na pasta onde deseja salvar o sistema e digite:
`git clone https://github.com/ifrass/extrator_planilha_analise.git`


## Como Usar no Dia a Dia
1. Abra a pasta do projeto e dê um clique duplo no arquivo **`iniciar.bat`**.
2. O sistema irá automaticamente procurar por atualizações na nuvem, instalar o que for necessário e abrir a interface no seu navegador web.
3. Arraste as planilhas da prefeitura para a área indicada.
4. Clique em **"Extrair Dados Limpos"**.
5. Baixe o arquivo gerado, selecione o bloco de dados e cole na sua Planilha Oficial usando **Colar Especial > Valores**.