# Extrator de Números de Planilhas Excel

Este aplicativo Streamlit permite processar qualquer aba de planilhas Excel e extrair números únicos com informações sobre consultores.

## Funcionalidades

- **Importar planilha Excel**: Upload de arquivos .xlsx
- **Seleção de aba**: Escolha qual aba da planilha processar
- **Extração de números**: Extrai números da coluna B (incluindo sequências com "-")
- **Criação de aba de resultados**: Gera nova aba com:
  - NUMERO: números únicos em ordem
  - ENCONTRADO_[ABA]: fórmula que verifica se o número existe
  - OCORRENCIAS_[ABA]: quantidade de consultores que contêm o número
  - CONSULTORES: lista de consultores onde o número aparece
- **Visualização dinâmica**: Interface para visualizar dados originais e processados
- **Download personalizado**: Baixar arquivo Excel modificado

## ✨ Novas Funcionalidades

- **Seleção flexível de abas**: Processe qualquer aba da planilha, não apenas REAT-10
- **Nomes personalizáveis**: Escolha o nome da aba de resultados
- **Interface dinâmica**: Títulos e visualizações se adaptam à aba selecionada
- **Tratamento de erros robusto**: Mensagens claras para problemas comuns
- **Compatibilidade total**: Funciona com qualquer estrutura de planilha Excel

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
   - Na barra lateral, clique em "Selecione o arquivo Excel"
   - Escolha sua planilha Excel (.xlsx)

3. **Selecionar aba para processar**:
   - Após o upload, escolha qual aba da planilha você quer processar
   - Personalize o nome da aba de saída (padrão: "EXTRAIDOS_REAT10")

4. **Visualizar dados originais**:
   - Na aba "📊 Dados Originais", você pode ver os dados da aba selecionada

5. **Processar dados**:
   - Clique no botão "Processar e atualizar aba"
   - Aguarde o processamento

6. **Visualizar resultados**:
   - Na aba "📈 Dados Processados", veja os dados extraídos e processados

7. **Baixar arquivo**:
   - Na aba "💾 Download", clique em "Baixar .xlsx atualizado"

## Estrutura esperada do arquivo Excel

O arquivo deve conter uma aba com dados estruturados da seguinte forma:
- **Coluna A**: Nomes dos consultores
- **Coluna B**: Sequências de números (podem conter hífens, ex: "123-456-789")

**Exemplos de abas que podem ser processadas:**
- REAT-10, VENDAS, DADOS_2024, RELATORIO, CONSULTORES, etc.
- Qualquer aba que tenha a estrutura: consultores na coluna A e números na coluna B

## Requisitos do sistema

- Python 3.7+
- Streamlit 1.28.0+
- Pandas 1.5.0+
- OpenPyXL 3.0.0+

## Solução de problemas

- **Erro "Arquivo não é um Excel válido"**: 
  - Verifique se o arquivo é realmente .xlsx (não .xls)
  - Tente salvar o arquivo novamente no Excel
  - Certifique-se de que o arquivo não está corrompido
- **Erro de dependências**: Execute `pip install -r requirements.txt` novamente
- **Problemas de memória**: Para arquivos muito grandes, considere dividir os dados em arquivos menores
- **Aba não encontrada**: Verifique se a aba selecionada existe na planilha
- **Dados não processados**: Certifique-se de que a aba tem consultores na coluna A e números na coluna B
