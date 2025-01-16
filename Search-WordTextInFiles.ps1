# Define uma classe para gerenciar a verificação de texto em arquivos Word
class WordTextSearch {
    [string]$SearchText
    [Object]$WordApp #[Microsoft.Office.Interop.Word]
    [string[]]$Extensions = @(".docx", ".docm", ".doc")

    # Construtor da classe, inicializando as variáveis
    WordTextSearch([string]$searchText) {
        $this.SearchText = $searchText
        
        # Inicializa a aplicação do Word
        $this.WordApp = New-Object -ComObject Word.Application
        $this.WordApp.Visible = $false
    }

    # Método para verificar se um documento contém o texto procurado
    [int]SearchTextInDocument([string]$filePath) {
        [int]$countOcurrences = 0
        #abre o documento no modo 'somente leitura'
        $document = $this.WordApp.Documents.Open($filePath, $false, $true) 
        # Memoriza ShowAll
        [bool]$ShowAllIsEnable = $document.ActiveWindow.ActivePane.View.ShowAll
        $document.ActiveWindow.ActivePane.View.ShowAll = $true
        #configura a consulta
        $finder = $this._setUpFinder([ref]$document)

        #inicia busca dentro do documento
        while ($finder.value.Find.Execute()) {
            ++$countOcurrences
        }

        #recupera o valor de ShowAll
        $document.ActiveWindow.ActivePane.View.ShowAll = $ShowAllIsEnable
        $document.Close([ref]0) #Fechar sem salvar

        return $countOcurrences
    }

    [System.Object]_setUpFinder([ref]$document){
        # Usar .Value para obter o objeto passado, já que é uma referência
        $finder = $this.WordApp.Selection
        $finder.Collapse(1)
        $finder.Find.ClearFormatting()
        $finder.Find.Text = $this.SearchText
        $finder.Find.Forward = $true
        $finder.Find.MatchCase = $false
        $finder.Find.Wrap = 0

        #retorna a referência
        return [ref]$finder
    }

    # Método para encerrar o aplicativo do Word
    [void]Dispose() {
        $this.WordApp.Quit()
    }
}

# Função PowerShell estilo cmdlet
function Search-WordTextInFiles {
    [CmdletBinding()]
    param (
        # Caminho da pasta onde estão os arquivos do Word; default para o diretório atual
        # -Path (1º Parametro)
        [Parameter(
            Position = 0,
            Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Path = (Get-Location),

        # Termo de pesquisa obrigatório
        # -SearchText (2º Parametro)
        [Parameter(
            Position = 1,
            Mandatory = $true
        )]
        [string]$SearchText
    )

    # Inicializa o objeto da classe com o termo de busca
    begin {
        $checker = [WordTextSearch]::new($SearchText)
    }

    # Processa cada item de entrada do pipeline
    process {
        # Verifica se $Path é um diretório
        if (-not (Test-Path -Path $Path -PathType Container)) {
            Write-Warning "Caminho invalido ou arquivo nao suportado: $Path"
            return
        }

        $files = Get-ChildItem -Path $Path | Where-Object { $_.Extension -in $checker.Extensions }
        $results = @()
        Write-Host "Aguarde. Buscando..."
        $files | ForEach-Object {
            $countOcurrences = $checker.SearchTextInDocument($_.FullName)

            if ($countOcurrences -gt 0) {
                $results += [PSCustomObject]@{
                    'Arquivo' = $_.Name
                    'Ocorrencias' = $countOcurrences
                }
            }
        }
        #### OUT ####
        $results
    }

    # Finaliza o cmdlet e libera o aplicativo do Word
    end {
        $checker.Dispose()
    }
}