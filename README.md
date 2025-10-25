# Extrator de Números da REAT-10

Este aplicativo Streamlit permite processar planilhas Excel com dados da REAT-10 e extrair números únicos com informações sobre consultores.

## Funcionalidades

- **Importar planilha Excel**: Upload de arquivos .xlsx
- **Processar aba REAT-10**: Extrai números da coluna B (incluindo sequências com "-")
- **Criar aba EXTRAIDOS_REAT10**: Gera nova aba com:
  - NUMERO: números únicos em ordem
  - ENCONTRADO_REAT10: fórmula que verifica se o número existe
  - OCORRENCIAS_REAT10: quantidade de consultores que contêm o número
  - CONSULTORES: lista de consultores onde o número aparece
- **Visualização**: Interface para visualizar os dados
- **Download**: Baixar arquivo Excel modificado

## Instalação

### Método 1: Usando o script automático (Windows)
1. **Execute o script**:
   ```bash
   run_app.bat
   ```

### Método 2: Instalação manual
1. **Criar ambiente virtual**:
   ```bash
   python -m venv venv
   ```

2. **Ativar ambiente virtual**:
   - **Windows**: `venv\Scripts\activate`
   - **Linux/Mac**: `source venv/bin/activate`

3. **Instalar dependências**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Executar o aplicativo**:
   ```bash
   streamlit run app.py
   ```

## Como usar

1. **Acesse o aplicativo**: O Streamlit abrirá automaticamente no navegador (geralmente em `http://localhost:8501`)

2. **Importar arquivo**: 
   - Na barra lateral, clique em "Selecione o arquivo .xlsx"
   - Escolha sua planilha Excel que contém a aba "REAT-10"

3. **Visualizar dados originais**:
   - Na aba "Visualizar REAT-10", você pode ver os dados originais

4. **Processar dados**:
   - Clique no botão "Processar e atualizar aba EXTRAIDOS_REAT10"
   - Aguarde o processamento

5. **Visualizar resultados**:
   - Na aba "EXTRAIDOS_REAT10 (visualização)", veja os dados processados

6. **Baixar arquivo**:
   - Na aba "Baixar arquivo modificado", clique em "Baixar .xlsx atualizado"

## Estrutura esperada do arquivo Excel

O arquivo deve conter uma aba chamada **"REAT-10"** com:
- **Coluna A**: Nomes dos consultores
- **Coluna B**: Sequências de números (podem conter hífens, ex: "123-456-789")

## Requisitos do sistema

- Python 3.7+
- Streamlit 1.28.0+
- Pandas 1.5.0+
- OpenPyXL 3.0.0+

## Solução de problemas

- **Erro "A aba 'REAT-10' não foi encontrada"**: Verifique se sua planilha tem uma aba exatamente com o nome "REAT-10"
- **Erro de dependências**: Execute `pip install -r requirements.txt` novamente
- **Problemas de memória**: Para arquivos muito grandes, considere dividir os dados em arquivos menores
