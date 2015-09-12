Private Sub Worksheet_Change(ByVal Target As Range)

    '**********  Este e o codigo principal que executa todos os eventos.  ***********

    Dim KeyCells As Range 'declara uma variavel que recebera dados de intervalo(Range)
    Set KeyCells = [DadosOrig] 'atribui a variavel o intervalo referente aos dados originais coletados na planilha.

    'Executa os eventos quando algum valor do intervalo for alterado:
    'Se algum dado do intervalo for alterado, entao...
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
        Application.ScreenUpdating = False 'desabilita a visualização do processo
        ActiveSheet.Unprotect '...Desprotege a planilha...
        '...executa os eventos abaixo...
        ClassificarDistribuirDados '(*)
        AjustarCasasDecimais '(**)
        ActiveSheet.Protect '...e protege a planilha novamente.
        [A1].Select 'seleciona a 1a celula da planilha.
    End If

End Sub

Private Sub ClassificarDistribuirDados() '(*)Classifica e distribui os dados na tabela de frequencia.

    Dim AreaDeClass As Range 'declara uma variavel que recebera dados de intervalo(Range)

    'Copia os valores da tabela original para outra area da planilha,
    'deixando os dados em uma unica coluna(U) para poder ordena-los em
    'ordem crescente. Copia uma linha de cada vez transpondo-a na coluna(U):
    [DadosOrig].Range("A1:H1").Copy 'copia primeira linha
    [U50].Offset(0, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=True 'transpoe em coluna
    [DadosOrig].Range("A2:H2").Copy 'copia segunda linha
    [U50].Offset(8, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=True 'transpoe em coluna logo abaixo da anterior
    [DadosOrig].Range("A3:H3").Copy 'copia terceira linha
    [U50].Offset(16, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=True 'transpoe em coluna logo abaixo da anterior
    [DadosOrig].Range("A4:H4").Copy 'copia quarta linha
    [U50].Offset(24, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=True 'transpoe em coluna logo abaixo da anterior

    'Copia a coluna(U) dos dados para outra coluna(AA) para posterior utilizacao (contagem de casas decimais):
    [U50:U100].Copy
    [AreaContagemCasasDec].PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

    'Remove os dados duplicados da coluna(U):
    ActiveSheet.[U50:U100].RemoveDuplicates Columns:=1, Header:=xlNo

    'Atribui o intervalo da coluna(U) a variavel inicialmente
    'declarada para ser a area de classificacao:
    Set AreaDeClass = [U50].Range("A1:A20") 'OBS: o intervalo tera no maximo 20 valores diferentes.

    'Classifica os dados da coluna(U) em ordem crescente:
    With Ex02.Sort
        .SortFields.Clear
        .SortFields.Add Key:=AreaDeClass, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
        .SetRange AreaDeClass
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Limpa os dados anteriores da tabela de frequencia, pois serao substituidos:
    [DadosClassif].ClearContents

    'Copia os dados classificados da coluna(U) para a tabela de frequencia:
    AreaDeClass.Copy
    [DadosClassif].Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    [U:U].ClearContents 'limpa os dados da coluna(U), nao ha mais necessidade deles

    [DadosClassif].EntireRow.Hidden = False 'reexibe as linhas ocultas da tabela de frequencia

    'Oculta as linhas nao utilizadas da tabela de frequencia
    'para nao aparecerem dados vazios no grafico:
    For i = 1 To 19
        [DadosClassif].Range("A1").End(xlDown).Offset(i, 0).EntireRow.Hidden = True
        If [DadosClassif].Range("A1").Offset(19, 0).EntireRow.Hidden = True Then
            Exit For
        End If
    Next i

End Sub

Sub AjustarCasasDecimais() '(**)Aplica a regra do Fix e ajusta as casas decimais na planilha.

    Dim i, j, x, Maior As Integer
    Dim vDados(50) As Double
    Dim vCasas(50) As Integer
    Dim FormatoNum As String

    'codigo para ajustar o valor de fix

    'atribui a variavel 'x' o valor da celula nomeada 'n' na planilha,
    'a qual representa o nro. de dados coletados. Decresce-se 'x' em 1 unidade
    'para efeito do laco que calcula qual o maior numero de casas decimais que ocorre nos dados originais
    x = [n].Value - 1
    Maior = 0 'antes de comecar o laco, ha zero casas decimais

    'varre cada dado original naquela coluna de dados anteriormente copiada
    'e verifica qual o maximo de casas decimais que ocorre entre eles.
    For i = 0 To x
        vDados(i) = [AreaContagemCasasDec].Range("A1").Offset(i, 0).Value
        For j = 1 To 6
            If Left(Right(vDados(i), j), 1) = "," Then
                vCasas(i) = j - 1
                If vCasas(i) > Maior Then
                    Maior = vCasas(i) 'a var. 'Maior' recebe a maior qtde de casas decimais ocorrido
                End If
            End If
        Next j
    Next i

    [AA:AA].ClearContents 'limpa os dados da area de contagem das casas decimais
    [AA50:AA100].Name = "AreaContagemCasasDec" 'recupera o nome do intervalo de celulas para proxima execucao

    'a celula que mostra o valor do Fix na planilha recebe o valor do maximo
    'de casas decimais ocorrido e acrescido em 1 unidade para aplicar a regra do Fix.
    [nCasasDec].Value = Maior + 1


    If Maior = 0 Then 'se nenhum dado original tiver casas decimais entao...
    FormatoNum = "0" '...a mascara do formato dos dados sera de numeros inteiros...
    [DadosOrig].NumberFormat = FormatoNum '...os dados originais recebem este formato...
    [Mo].NumberFormat = FormatoNum '...o valor da 'Moda' recebe este formato...
    [Min].NumberFormat = FormatoNum '...o valor do 'Min' recebe este formato...
    [Max].NumberFormat = FormatoNum '...o valor do 'Max' recebe este formato...
    [DadosClassif].NumberFormat = FormatoNum '...e os dados classificados na tabela de frequencia também.
    'por fim, acrescenta-se na mascara uma casa decimal para usar a regra do fix nos outros valores:
    FormatoNum = FormatoNum & ".0"
    Else 'senao se ao menos 1 dado possuir casa decimal...
    FormatoNum = "0." '...a mascara inicia com o prefixo zero-virgula para posterior acrescimo de zeros.
    '...o laco vai de 1 casa ate o valor de Fix calculado anteriormente para acrescentar os zeros e...
    For i = 1 To [nCasasDec].Value
        [DadosOrig].NumberFormat = FormatoNum '... os dados originais recebem o formato
        [Mo].NumberFormat = FormatoNum '...o valor da 'Moda' recebe o formato...
        [Min].NumberFormat = FormatoNum '...o valor do 'Min' recebe o formato...
        [Max].NumberFormat = FormatoNum '...o valor do 'Max' recebe o formato...
        [DadosClassif].NumberFormat = FormatoNum '...e os dados da tabela de frequencia tambem.
        'por fim, acrescenta-se na mascara uma casa decimal para usar a regra do fix nos outros valores:
        FormatoNum = FormatoNum & "0"
    Next i
End If

'os intervalos abaixo, que sao os valores calculados na planilha,
'recebem os formatos adequados a regra do Fix:
[xifi].NumberFormat = FormatoNum
[xi2fi].NumberFormat = FormatoNum
[Media].NumberFormat = FormatoNum
[Quartis].NumberFormat = FormatoNum
[Percentis].NumberFormat = FormatoNum
[s].NumberFormat = FormatoNum
[CC].NumberFormat = FormatoNum
[CoefVar].NumberFormat = FormatoNum
[CoefAssim].NumberFormat = FormatoNum
[Col_xifi].NumberFormat = FormatoNum
[Col_xi2fi].NumberFormat = FormatoNum

End Sub