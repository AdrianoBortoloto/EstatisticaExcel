Option Explicit
Dim i, Maior, j As Integer
Dim KeyCells As Range
Dim vDados(20) As Double
Dim vCasas(20) As Integer
Dim FormatoNum As String

Private Sub Worksheet_Change(ByVal Target As Range)

    Set KeyCells = [Dados]
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
        Application.ScreenUpdating = False
        ActiveSheet.Unprotect
        AjustarCasasDecimais
    End If

End Sub

Sub BotaoAlterarAdicRemov()

    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    [ColX].EntireRow.Hidden = False
    ActiveSheet.Protect

End Sub

Sub BotaoOcultarLinhasVazias()

    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    For i = 1 To 19
        If [ColX].Range("A1").Offset(i, 0).Value = "" Then
            [ColX].Range("A1").Offset(i, 0).EntireRow.Hidden = True
        End If
    Next i
    ActiveSheet.Protect

End Sub

Private Sub AjustarCasasDecimais()

    'codigo para ajustar o valor de fix
    Maior = 0
    For i = 0 To [n].Value - 1
        vDados(i) = [ColX].Range("A1").Offset(i, 0).Value
        For j = 1 To 6
            If Left(Right(vDados(i), j), 1) = "," Then
                vCasas(i) = j - 1
                If vCasas(i) > Maior Then
                    Maior = vCasas(i)
                End If
            End If
        Next j
    Next i

    For i = 0 To [n].Value - 1
        vDados(i) = [ColY].Range("A1").Offset(i, 0).Value
        For j = 1 To 6
            If Left(Right(vDados(i), j), 1) = "," Then
                vCasas(i) = j - 1
                If vCasas(i) > Maior Then
                    Maior = vCasas(i)
                End If
            End If
        Next j
    Next i


    [nCasas].Value = Maior + 1

    If Maior = 0 Then
        FormatoNum = "0"
        [ColX].NumberFormat = FormatoNum
        [ColY].NumberFormat = FormatoNum
        FormatoNum = FormatoNum & ".0"
    Else
        FormatoNum = "0."
        For i = 1 To [nCasas].Value
            [ColX].NumberFormat = FormatoNum
            [ColY].NumberFormat = FormatoNum
            FormatoNum = FormatoNum & "0"
        Next i
    End If

    'Altera o numero de casas apos a virgula (Fix) dos intervalos previamente nomeados na planilha.
    [XY].NumberFormat = FormatoNum 'variavel "FormatoNum" mencionada em comentario anterior.
    [XQuad].NumberFormat = FormatoNum
    [YQuad].NumberFormat = FormatoNum
    [Sxy].NumberFormat = FormatoNum
    [Sx].NumberFormat = FormatoNum
    [Sy].NumberFormat = FormatoNum
    [a].NumberFormat = FormatoNum
    [x_linha].NumberFormat = FormatoNum
    [y_linha].NumberFormat = FormatoNum
    [b].NumberFormat = FormatoNum
    [y_Estimativa].NumberFormat = FormatoNum
    [y_Estimativa].NumberFormat = "General ""ms"""

    ActiveSheet.Protect

End Sub