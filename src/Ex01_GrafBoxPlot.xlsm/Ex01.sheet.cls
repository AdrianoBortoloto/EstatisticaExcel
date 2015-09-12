Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range
    Set KeyCells = Range("DadosOrig")

    'executa os eventos quando a(s) celula(s) atribuida a variavel KeyCells for(em) alterada(s)
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
        Application.ScreenUpdating = False          'desabilita a visualizacao do processo
        ActiveSheet.Unprotect                       'desprotege a planilha
        ClassifDadosOrdemCrescente                  'ordena os dados em ordem crescente na tabela
        AjustaNumCasasDecimais                      'ajusta a precisao conforme a regra do Fix
        ActiveSheet.Protect                         'protege a planilha
        Range(Target.Address).Offset(1, 0).Select   'seleciona a próxima linha
    End If

End Sub

Private Sub ClassifDadosOrdemCrescente()

    Dim AreaDeClassif As String

    'area onde os dados sao copiados temporariamente para ordena-los em ordem crescente
    AreaDeClassif = "S3:S14"

    [DadosOrig].Copy 'copia os dados da tabela original para memoria
    [S3].PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False 'cola na area de classificacao

    'o bloco abaixo trata os parametros de ordenacao
    With Ex01.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range(AreaDeClassif), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
        .SetRange Range(AreaDeClassif)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range(AreaDeClassif).Copy 'copia os dados ja ordenados para memoria
    [E3].PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=True 'cola os dados na tabela dos dados classificados

    Range(AreaDeClassif).ClearContents 'limpa o conteudo da area de classificacao
    '[A1].Select

End Sub

Private Sub AjustaNumCasasDecimais()

    Dim i, j, x, Maior As Integer
    Dim vDados(15) As Double 'vetor para os dados
    Dim vCasas(15) As Integer 'vetor de saida para o numero de casas decimais de cada dado
    Dim FormatoNum, aux As String

    'o codigo abaixo faz a varredura em cada dado procurando a virgula da
    'direita para a esquerda contando quantas casas possui ate encontrar a virgula

    x = [n].Value - 1 ' "n" e o nome da celula que contem a qtde de dados
    Maior = 0
    For i = 0 To x
        vDados(i) = [DadosOrig].Range("A1").Offset(i, 0).Value
        For j = 1 To 6  'o nro. 6 limita a varredura ate seis casas apos a virgula
            If Left(Right(vDados(i), j), 1) = "," Then  'se o laco encontrar virgula entao...
            vCasas(i) = j - 1 '...o numero de casas e igual ao numero de loops - 1
            If vCasas(i) > Maior Then 'se o numero de casas for maior que o anterior entao...
            Maior = vCasas(i) '...o maior passa a ser o valor atual
        End If
    End If
Next j
    Next i
        
    'apos achar qual e o dado que possui mais casas decimais,
    'a celula na planilha nomeada "nCasas" e preenchida com o valor do Fix.
    [nCasas].Value = Maior + 1

    If Maior = 0 Then       'se nenhum dado possuir casas decimais entao...
        FormatoNum = "0"                            '...a mascara do formato do numero sera sem virgula...
        [DadosOrig].NumberFormat = FormatoNum       '...os dados originais tomam o formato da mascara...
        [Mo].NumberFormat = FormatoNum              '...a celula referente a moda toma o mesmo formato...
        [DadosClass].NumberFormat = FormatoNum      ' ...e os dados classificados tambem.
        FormatoNum = FormatoNum & ".0"              '... e por fim, ajusta o formato conforme o valor do Fix.
    Else    'senao caso ao menos 1 dado tiver casa decimal, executa o laco abaixo para ajustar...
        FormatoNum = "0."                           '...a mascara inicial para aplicar no laco 'For' abaixo...
        For i = 1 To Range("nCasas").Value          '...fara tantos loops quanto nro. de casas...
            [DadosOrig].NumberFormat = FormatoNum   '...os dados originais tomam o formato da mascara...
            [Mo].NumberFormat = FormatoNum          '...a celula referente a moda toma o mesmo formato...
            [DadosClass].NumberFormat = FormatoNum  '...e os dados classificados tambem.
            FormatoNum = FormatoNum & "0"           'a cada loop ,acrescenta um zero apos a virgula.
        Next i
    End If
    
    'o resultado da mascara e aplicado em todos os valores necessarios da planilha
    [DadosClassif2].NumberFormat = FormatoNum
    [Somat_xi].NumberFormat = FormatoNum
    [Somat_xi2].NumberFormat = FormatoNum
    [Media].NumberFormat = FormatoNum
    [Separatrizes].NumberFormat = FormatoNum
    [DesvPad].NumberFormat = FormatoNum
    [CoefVar].NumberFormat = FormatoNum & "%"
    [CoefAssim].NumberFormat = FormatoNum
    [Dq].NumberFormat = FormatoNum
    [L.Sup].NumberFormat = FormatoNum
    [L.Inf].NumberFormat = FormatoNum
    [DadosGrafico].NumberFormat = FormatoNum
    
End Sub

