Option Explicit
Dim KeyCells As Range
Dim FormatoNum As String
Dim a, i, j, nLTab, nCTab, nClasses, conta, Maior As Integer
Dim vDados(55) As Double
Dim vCasas(55) As Integer

Private Sub Worksheet_Change(ByVal Target As Range)

    Set KeyCells = [DadosOrig]
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
        Application.ScreenUpdating = False
        ActiveSheet.Unprotect
        DistribuirFrequencias
        AjustarCasasDecimais
        ActiveSheet.Protect
        [A1].Select
    End If

End Sub
Sub DistribuirFrequencias()

    'distribui as frequencias dos valores nas classes
    nLTab = [nLin].Value 'numero de linhas da tabela de dados contido na celula nomeada "nLin".
    nCTab = [nCol].Value 'numero de colunas da tabela de dados contido na celula nomeada "nCol".
    nClasses = [k].Value 'valor de k (numero de classes) contido na celula nomeada "k"

    [Tab_fi].ClearContents
    For a = 1 To nClasses
        For i = 1 To nLTab
            For j = 1 To nCTab
                If a = nClasses Then ' se o laco estiver no valor maximo, ou seja, na ultima classe,
                ' o valor devera ser < ou = ao Limite superior(L) da classe(k).
                If [DadosOrig].Cells(i, j).Value >= [Intervalos].Cells(a, 1).Value And _
                    [DadosOrig].Cells(i, j).Value <= [Intervalos].Cells(a, 2).Value Then
                    conta = conta + 1 ' acumula o valor da frequencia da ultima classe.
                End If
            Else
                If [DadosOrig].Cells(i, j).Value >= [Intervalos].Cells(a, 1).Value And _
                    [DadosOrig].Cells(i, j).Value < [Intervalos].Cells(a, 2).Value Then
                    conta = conta + 1 ' acumula o valor da frequencia das classes, exceto a ultima.
                End If
            End If
        Next j
    Next i

    'mostra o valor da frequencia na celula da classe correspondente
    [Intervalos].Cells(a, 4).Value = conta
    conta = 0 ' aqui a contagem e zerada para fazer o acumulo da frequencia da proxima classe.
Next a

'Oculta as linhas nao utilizadas da tabela de frequencia
'para nao aparecerem dados vazios no grafico:
[ColClasses].EntireRow.Hidden = False
For i = 1 To 19
    If [ColClasses].Range("A1").Offset(i, 0).Value = "" Then
        [ColClasses].Range("A1").Offset(i, 0).EntireRow.Hidden = True
    End If
Next i

End Sub

Sub AjustarCasasDecimais()
    
    'Copia os valores da tabela original para outra area da planilha,
    'deixando os dados em uma unica coluna(U). Copia uma linha de cada
    'vez, transpondo-a em coluna:
    [DadosOrig].Range("A1:K1").Copy 'copia primeira linha
    [U50].Offset(0, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True 'transpoe em coluna
    [DadosOrig].Range("A2:K2").Copy 'copia segunda linha
    [U50].Offset(11, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True 'transpoe em coluna logo abaixo da anterior
    [DadosOrig].Range("A3:K3").Copy 'copia terceira linha
    [U50].Offset(22, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True 'transpoe em coluna logo abaixo da anterior
    [DadosOrig].Range("A4:K4").Copy 'copia quarta linha
    [U50].Offset(33, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True 'transpoe em coluna logo abaixo da anterior
    [DadosOrig].Range("A5:K5").Copy 'copia quinta linha
    [U50].Offset(44, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True 'transpoe em coluna logo abaixo da anterior

    'codigo para ajustar o valor de fix
    Maior = 0
    For i = 0 To [n].Value - 1
        vDados(i) = [AreaContagemCasasDec].Range("A1").Offset(i, 0).Value
        For j = 1 To 6
            If Left(Right(vDados(i), j), 1) = "," Then
                vCasas(i) = j - 1
                If vCasas(i) > Maior Then
                    Maior = vCasas(i)
                End If
            End If
        Next j
    Next i

    [AreaContagemCasasDec].ClearContents

    [nCasas].Value = Maior + 1

    If Maior = 0 Then
        FormatoNum = "0"
        [DadosOrig].NumberFormat = FormatoNum
        FormatoNum = FormatoNum & ".0"
    Else
        FormatoNum = "0."
        For i = 1 To [nCasas].Value
            [DadosOrig].NumberFormat = FormatoNum
            FormatoNum = FormatoNum & "0"
        Next i
    End If

    'Altera o numero de casas apos a virgula (Fix) dos intervalos previamente nomeados na planilha.
    [AA].NumberFormat = FormatoNum 'variavel "FormatoNum" mencionada em comentario anterior.
    [h].NumberFormat = FormatoNum
    [Media].NumberFormat = FormatoNum
    [Mo].NumberFormat = FormatoNum
    [Intervalos].NumberFormat = FormatoNum
    [Col_xi].NumberFormat = FormatoNum
    [Cols_xifi_xi2fi].NumberFormat = FormatoNum
    [Quartis].NumberFormat = FormatoNum
    [Percentis].NumberFormat = FormatoNum
    [s].NumberFormat = FormatoNum
    [CoefVar].NumberFormat = FormatoNum & "%"
    [CoefAssim].NumberFormat = FormatoNum
    [CC].NumberFormat = FormatoNum

End Sub
