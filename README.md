# Storni_Bind
Project to viably create excel addin that uses pythons
Guia de Instalação e Configuração do Add-in Excel-Python
Pré-requisitos
1. Python e Bibliotecas
Certifique-se de ter o Python instalado com as seguintes bibliotecas:

bash
pip install pandas numpy openpyxl xlrd
2. Verificar Caminho do Python
Abra o prompt de comando e digite:

bash
python --version
where python
Anote o caminho completo para usar na configuração.

Instalação do Add-in
Passo 1: Criar os Arquivos
Crie uma pasta para o add-in (ex: C:\ExcelPythonAddin\)
Salve o arquivo excel_functions.py nesta pasta
Ajuste o caminho no código VBA conforme necessário
Passo 2: Configurar o VBA no Excel
Abra o Excel
Pressione Alt + F11 para abrir o Editor VBA
No menu Insert > Module
Cole o código VBA fornecido
Ajuste as constantes no início do código:
vba
Private Const PYTHON_PATH As String = "C:\Python39\python.exe" ' Seu caminho
Private Const SCRIPT_PATH As String = "C:\ExcelPythonAddin\excel_functions.py"
Passo 3: Salvar como Add-in
No Excel, vá em File > Save As
Escolha o tipo "Excel Add-in (*.xlam)"
Salve com nome descritivo (ex: "PythonDataProcessor.xlam")
Salve na pasta de add-ins do Excel (geralmente: %APPDATA%\Microsoft\AddIns\)
Passo 4: Ativar o Add-in
No Excel, vá em File > Options > Add-ins
Na parte inferior, selecione "Excel Add-ins" e clique "Go"
Marque seu add-in na lista
Clique "OK"
Como Usar
Função Principal: ProcessExcelData
vba
=ProcessExcelData(A1:D10) ' Processa range e salva com nome automático
=ProcessExcelData(A1:D10,"C:\caminho\meu_arquivo.txt") ' Especifica arquivo de saída
Função de Análise: AnalyzeExcelData
vba
=AnalyzeExcelData(A1:D10) ' Retorna estatísticas descritivas dos dados
Usando no VBA (mais avançado)
vba
Sub ProcessarDados()
    Dim resultado As String
    resultado = ProcessExcelData(Range("A1:D10"), "C:\output\dados.txt")
    MsgBox resultado
End Sub
Personalização das Funções Python
Edite o arquivo excel_functions.py para adicionar suas próprias funções:

python
def minha_funcao_personalizada(data_range, parametro1, parametro2):
    """
    Sua função personalizada aqui
    """
    df = pd.DataFrame(data_range)
    
    # Seu código de processamento aqui
    
    return resultado
Estrutura de Arquivos Recomendada
C:\ExcelPythonAddin\
├── excel_functions.py          # Funções Python
├── config.ini                  # Configurações (opcional)
├── logs\                       # Pasta para logs (opcional)
└── output\                     # Pasta padrão para saídas
Solução de Problemas
Erro: "Python não encontrado"
Verifique se o caminho do Python está correto no código VBA
Teste o comando python no prompt de comando
Erro: "Módulo pandas não encontrado"
Execute: pip install pandas numpy
Se usar Anaconda: conda install pandas numpy
Erro: "Permissão negada"
Execute o Excel como administrador
Verifique permissões da pasta de destino
Erro: "Macro desabilitada"
Vá em File > Options > Trust Center > Trust Center Settings
Em "Macro Settings", selecione "Enable VBA macros"
Funcionalidades Avançadas
1. Interface com Botões
Para criar botões personalizados, adicione ao VBA:

vba
Sub AddCustomButtons()
    ' Código para adicionar botões à ribbon do Excel
End Sub
2. Processamento Assíncrono
Para grandes volumes de dados, considere implementar processamento em background.

3. Cache de Resultados
Implemente cache para evitar reprocessamento desnecessário de dados.

Exemplo de Uso Prático
Selecione uma tabela no Excel (ex: A1:E100)
Em uma célula vazia, digite: =ProcessExcelData(A1:E100,"C:\dados\resultado.txt")
O arquivo será processado pelo pandas e salvo como TXT
A função retornará uma mensagem confirmando o salvamento
Notas Importantes
O add-in funciona melhor com dados estruturados (tabelas)
Certifique-se de ter permissões de escrita nos diretórios de saída
Para grandes volumes de dados, considere usar processamento em lotes
Mantenha backups dos seus arquivos Python personalizados
