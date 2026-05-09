Attribute VB_Name = "Módulo1"
Option Explicit
Public SENHA As String
Public ProximoAutoSave As Double
Public dictMateriais As Object ' cache em memória (código > descrição)

Public Sub OtimizarON()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub OtimizarOFF()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub


' ==============================================
' FUNÇÃO: ProximoLote
' ==============================================
Function ProximoLote() As String

    On Error GoTo ErroLote

    Dim wsConf As Worksheet
    Set wsConf = ThisWorkbook.Sheets("Configurações")

    Dim prox As Long
    Dim valorCelula As String

    valorCelula = Trim(CStr(wsConf.Range("B5").Value))

    If valorCelula = "" Then
        
        prox = 1
        
    ElseIf IsNumeric(valorCelula) Then
        
        prox = CLng(valorCelula)
        
    Else
        
        MsgBox "Valor inválido em Configurações! B5 precisa ser numérico.", vbCritical
        prox = 1
        
    End If

    wsConf.Range("B5").Value = prox + 1

    ProximoLote = Format(prox, "0000000000")

    Exit Function

ErroLote:

    MsgBox _
        "Erro ao gerar lote:" & vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical

    ProximoLote = ""

End Function


' ==============================================
' ENTRADA
' ==============================================
Sub RegistrarEntrada()

    On Error GoTo Finalizar

    ' =========================
    ' SEGURANÇA + PERFORMANCE
    ' =========================
    If SENHA = "" Then SENHA = GetSenha()
    
    Call OtimizarON
    Call DesprotegerSistema

    Dim wsEntrada As Worksheet, wsEstoque As Worksheet, wsHistorico As Worksheet
    Set wsEntrada = ThisWorkbook.Sheets("Entradas")
    Set wsEstoque = ThisWorkbook.Sheets("Estoque Atual")
    Set wsHistorico = ThisWorkbook.Sheets("Historico")

    Dim ultimaLinha As Long
    ultimaLinha = wsEntrada.Cells(wsEntrada.Rows.Count, "A").End(xlUp).Row

    Dim temDados As Boolean
    Dim x As Long

    ' =========================
    ' VERIFICA SE TEM DADOS
    ' =========================
    For x = 2 To ultimaLinha
        If Trim(wsEntrada.Cells(x, 1).Value) <> "" Then
            temDados = True
            Exit For
        End If
    Next x

    If Not temDados Then
        MsgBox "Nenhuma linha preenchida!", vbExclamation
        GoTo Finalizar
    End If

    Dim i As Long
    Dim erros As String
    Dim temErro As Boolean

    ' =========================
    ' FASE 1 - VALIDAÇÃO
    ' =========================
    For i = 2 To ultimaLinha

        Dim codMat As String
        Dim notaFiscal As String
        Dim qty As Double
        Dim descMat As String

        codMat = Trim(wsEntrada.Cells(i, 1).Value)
        notaFiscal = Trim(wsEntrada.Cells(i, 4).Value)
        descMat = BuscarDescricaoRapida(codMat)

        If codMat = "" And notaFiscal = "" And wsEntrada.Cells(i, 3).Value = "" Then GoTo ProximaValidacao

        wsEntrada.Range("A" & i & ":E" & i).Interior.ColorIndex = xlNone

        If codMat = "" Then
            erros = erros & "Linha " & i & ": Código vazio" & vbCrLf
            wsEntrada.Range("A" & i & ":E" & i).Interior.Color = RGB(255, 200, 200)
            temErro = True
            GoTo ProximaValidacao
        End If

        If Not IsNumeric(wsEntrada.Cells(i, 3).Value) Then
            erros = erros & "Linha " & i & ": Quantidade inválida" & vbCrLf
            wsEntrada.Range("A" & i & ":E" & i).Interior.Color = RGB(255, 200, 200)
            temErro = True
            GoTo ProximaValidacao
        End If

        qty = CDbl(wsEntrada.Cells(i, 3).Value)

        If qty <= 0 Then
            erros = erros & "Linha " & i & ": Quantidade <= 0" & vbCrLf
            wsEntrada.Range("A" & i & ":E" & i).Interior.Color = RGB(255, 200, 200)
            temErro = True
            GoTo ProximaValidacao
        End If

        If notaFiscal = "" Then
            erros = erros & "Linha " & i & ": Nota Fiscal obrigatória" & vbCrLf
            wsEntrada.Range("A" & i & ":E" & i).Interior.Color = RGB(255, 200, 200)
            temErro = True
            GoTo ProximaValidacao
        End If

        If descMat = "[NÃO CADASTRADO]" Then
            erros = erros & "Linha " & i & ": Material não cadastrado" & vbCrLf
            wsEntrada.Range("A" & i & ":E" & i).Interior.Color = RGB(255, 200, 200)
            temErro = True
            GoTo ProximaValidacao
        End If

ProximaValidacao:
    Next i

    If temErro Then
        MsgBox "Erros encontrados:" & vbCrLf & vbCrLf & erros, vbExclamation
        Application.GoTo wsEntrada.Range("A2"), True
        GoTo Finalizar
    End If

    ' =========================
    ' FASE 2 - PROCESSAMENTO
    ' =========================
    For i = 2 To ultimaLinha

        codMat = Trim(wsEntrada.Cells(i, 1).Value)

        If codMat = "" Then GoTo ProximaExecucao

        notaFiscal = wsEntrada.Cells(i, 4).Value
        qty = CDbl(wsEntrada.Cells(i, 3).Value)
        descMat = BuscarDescricaoRapida(codMat)

        Dim lote As String
        lote = ProximoLote()
        If lote = "" Then
            MsgBox "Falha ao gerar lote.", vbCritical
            GoTo Finalizar
        End If

        Dim linhaEstoque As Long
        linhaEstoque = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row + 1

        ' ESTOQUE
        wsEstoque.Cells(linhaEstoque, 1).Value = codMat
        wsEstoque.Cells(linhaEstoque, 2).Value = descMat
        wsEstoque.Cells(linhaEstoque, 3).Value = lote
        wsEstoque.Cells(linhaEstoque, 4).Value = qty
        wsEstoque.Cells(linhaEstoque, 5).Value = 0
        wsEstoque.Cells(linhaEstoque, 6).Value = 0
        wsEstoque.Cells(linhaEstoque, 8).Value = notaFiscal

        ' HISTÓRICO
        Dim h As Long
        h = wsHistorico.Cells(wsHistorico.Rows.Count, "A").End(xlUp).Row + 1

        wsHistorico.Cells(h, 1).Value = codMat
        wsHistorico.Cells(h, 2).Value = descMat
        wsHistorico.Cells(h, 3).Value = lote
        wsHistorico.Cells(h, 4).Value = qty
        wsHistorico.Cells(h, 5).Value = "Entrada"
        wsHistorico.Cells(h, 6).Value = Environ("USERNAME")
        wsHistorico.Cells(h, 7).Value = Now()
        wsHistorico.Cells(h, 8).Value = notaFiscal
        wsHistorico.Cells(h, 9).Value = ""
        wsHistorico.Cells(h, 10).Value = ""

        Call ColorirLinhaHistorico(wsHistorico, h, "Entrada")

ProximaExecucao:
    Next i

    ' =========================
    ' LIMPA TUDO
    ' =========================
    wsEntrada.Range("A2:D" & ultimaLinha).ClearContents
    wsEntrada.Range("A2:D" & ultimaLinha).Interior.ColorIndex = xlNone

    Call OrdenarEstoque
    Call FiltrarEstoque
    Call OrdenarHistorico
    Call SalvarSeguro
    

    Application.Calculate

    MsgBox "Entradas processadas com sucesso!", vbInformation
    Application.GoTo wsEntrada.Range("A2"), True

Finalizar:
    If Err.Number <> 0 Then

    MsgBox _
        "Erro " & Err.Number & vbCrLf & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Sub: RegistrarEntrada", _
        vbCritical

    End If
    
    Call ProtegerSistema
    Call OtimizarOFF

End Sub

' ==============================================
' BAIXA
' ==============================================
Sub DarBaixa()

On Error GoTo Finalizar

If SENHA = "" Then SENHA = GetSenha()
Call DesprotegerSistema

Dim wsBaixa As Worksheet, wsEstoque As Worksheet, wsHistorico As Worksheet
Set wsBaixa = ThisWorkbook.Sheets("Baixas")
Set wsEstoque = ThisWorkbook.Sheets("Estoque Atual")
Set wsHistorico = ThisWorkbook.Sheets("Historico")

Dim ultLinha As Long
Dim i As Long, j As Long
Dim codMat As String, lote As String
Dim qty As Double
Dim ultEstoque As Long
Dim encontrado As Boolean
Dim erros As String
Dim temErro As Boolean
Dim temDados As Boolean

Call OtimizarON

ultLinha = wsBaixa.Cells(wsBaixa.Rows.Count, "A").End(xlUp).Row

' valida dados
For i = 2 To ultLinha
    If Trim(wsBaixa.Cells(i, 1).Value) <> "" Then
        temDados = True
        Exit For
    End If
Next i

If Not temDados Then
    MsgBox "Nenhuma linha preenchida!", vbExclamation
    GoTo Finalizar
End If

' limpa cores
For i = 2 To ultLinha
    wsBaixa.Range("A" & i & ":D" & i).Interior.ColorIndex = xlNone
Next i

' validação
For i = 2 To ultLinha

    codMat = Trim(wsBaixa.Cells(i, 1).Value)
    lote = Trim(wsBaixa.Cells(i, 4).Value)

    If codMat = "" And lote = "" Then GoTo ProximaValidacao

    If codMat = "" Or lote = "" Then
        erros = erros & "Linha " & i & ": Código ou lote vazio" & vbCrLf
        wsBaixa.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)
        temErro = True
        GoTo ProximaValidacao
    End If

    If Not IsNumeric(wsBaixa.Cells(i, 3).Value) Then
        erros = erros & "Linha " & i & ": Quantidade inválida" & vbCrLf
        wsBaixa.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)
        temErro = True
        GoTo ProximaValidacao
    End If

    qty = CDbl(wsBaixa.Cells(i, 3).Value)

    ultEstoque = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row
    encontrado = False

    For j = ultEstoque To 2 Step -1

        If wsEstoque.Cells(j, 1).Value = codMat And wsEstoque.Cells(j, 3).Value = lote Then

            encontrado = True

            If qty > wsEstoque.Cells(j, 4).Value Then
                erros = erros & "Linha " & i & ": Saldo insuficiente" & vbCrLf
                wsBaixa.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)
                temErro = True
            End If

            Exit For
        End If

    Next j

    If Not encontrado Then
        erros = erros & "Linha " & i & ": Lote não encontrado" & vbCrLf
        wsBaixa.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)
        temErro = True
    End If

ProximaValidacao: Next i

If temErro Then
    MsgBox "Erros encontrados:" & vbCrLf & vbCrLf & erros, vbExclamation
    GoTo Finalizar
End If

' processamento
For i = 2 To ultLinha

    codMat = Trim(wsBaixa.Cells(i, 1).Value)
    lote = Trim(wsBaixa.Cells(i, 4).Value)

    If codMat = "" And lote = "" Then GoTo ProximaLinha

    qty = CDbl(wsBaixa.Cells(i, 3).Value)

    ultEstoque = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

    For j = ultEstoque To 2 Step -1

        If wsEstoque.Cells(j, 1).Value = codMat And wsEstoque.Cells(j, 3).Value = lote Then

            Dim saldo As Double
            saldo = wsEstoque.Cells(j, 4).Value

            Dim desc As String
            desc = wsEstoque.Cells(j, 2).Value

            Dim notaFiscal As String
            notaFiscal = wsEstoque.Cells(j, 8).Value

            wsEstoque.Cells(j, 4).Value = saldo - qty

            Dim livre As Double, qualidade As Double, bloqueado As Double
            livre = wsEstoque.Cells(j, 4).Value
            qualidade = wsEstoque.Cells(j, 5).Value
            bloqueado = wsEstoque.Cells(j, 6).Value

            Dim h As Long
            h = wsHistorico.Cells(wsHistorico.Rows.Count, "A").End(xlUp).Row + 1

            wsHistorico.Cells(h, 1).Value = codMat
            wsHistorico.Cells(h, 2).Value = desc
            wsHistorico.Cells(h, 3).Value = lote
            wsHistorico.Cells(h, 4).Value = qty
            wsHistorico.Cells(h, 5).Value = "Baixa"
            wsHistorico.Cells(h, 6).Value = Environ("USERNAME")
            wsHistorico.Cells(h, 7).Value = Now()
            wsHistorico.Cells(h, 8).Value = notaFiscal

            Call ColorirLinhaHistorico(wsHistorico, h, "Baixa")

            If livre + qualidade + bloqueado = 0 Then
                wsEstoque.Rows(j).Delete
            End If

            Exit For
        End If

    Next j

ProximaLinha:

    wsBaixa.Range("A" & i & ":D" & i).ClearContents
    wsBaixa.Range("A" & i & ":D" & i).Interior.ColorIndex = xlNone

Next i

Call OrdenarEstoque
Call FiltrarEstoque
Call OrdenarHistorico
Call SalvarSeguro


MsgBox "Baixas processadas com sucesso!", vbInformation

Finalizar:

If Err.Number <> 0 Then

MsgBox _
    "Erro " & Err.Number & vbCrLf & vbCrLf & _
    Err.Description & vbCrLf & vbCrLf & _
    "Sub: DarBaixa", _
    vbCritical

End If

Call ProtegerSistema

Call OtimizarOFF

End Sub
' ==============================================
' BLOQUEAR
' ==============================================
Sub BloquearMaterial()

    On Error GoTo Finalizar

    If SENHA = "" Then SENHA = GetSenha()
    Call DesprotegerSistema

    Dim ws As Worksheet
    Dim wsEstoque As Worksheet
    Dim wsHist As Worksheet

    Set ws = ThisWorkbook.Sheets("Bloqueados")
    Set wsEstoque = ThisWorkbook.Sheets("Estoque Atual")
    Set wsHist = ThisWorkbook.Sheets("Historico")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim ultLinha As Long
    ultLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim temDados As Boolean
    Dim x As Long

    ' =========================
    ' VERIFICA DADOS
    ' =========================
    For x = 2 To ultLinha

        If Trim(ws.Cells(x, 1).Value) <> "" Then
            temDados = True
            Exit For
        End If

    Next x

    If Not temDados Then
        MsgBox "Nenhuma linha preenchida!", vbExclamation
        GoTo Finalizar
    End If

    Dim erros As String
    Dim temErro As Boolean
    Dim i As Long

    ' =========================
    ' LIMPA CORES ANTIGAS
    ' =========================
    For i = 2 To ultLinha
        ws.Range("A" & i & ":D" & i).Interior.ColorIndex = xlNone
    Next i

    ' =========================
    ' FASE 1 - VALIDAÇÃO
    ' =========================
    For i = 2 To ultLinha

        Dim cod As String
        Dim lote As String
        Dim qty As Double
        Dim encontrado As Boolean

        encontrado = False

        cod = Trim(ws.Cells(i, 1).Value)
        lote = Trim(ws.Cells(i, 4).Value)

        If cod = "" And lote = "" Then GoTo ProxValidacao

        If cod = "" Or lote = "" Then

            erros = erros & "Linha " & i & ": Código ou lote vazio" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValidacao

        End If

        If Not IsNumeric(ws.Cells(i, 3).Value) Then

            erros = erros & "Linha " & i & ": Quantidade inválida" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValidacao

        End If

        qty = CDbl(ws.Cells(i, 3).Value)

        If qty <= 0 Then

            erros = erros & "Linha " & i & ": Quantidade <= 0" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValidacao

        End If

        Dim j As Long
        Dim ultEstoque As Long

        ultEstoque = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

        For j = 2 To ultEstoque

            If wsEstoque.Cells(j, 1).Value = cod And _
               wsEstoque.Cells(j, 3).Value = lote Then

                encontrado = True

                Dim livre As Double
                livre = Val(wsEstoque.Cells(j, 4).Value)

                If qty > livre Then

                    erros = erros & "Linha " & i & ": Saldo insuficiente" & vbCrLf

                    ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

                    temErro = True

                End If

                Exit For

            End If

        Next j

        If Not encontrado Then

            erros = erros & "Linha " & i & ": Lote não encontrado" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

        End If

ProxValidacao:
    Next i

    ' =========================
    ' BLOQUEIA PROCESSAMENTO
    ' =========================
    If temErro Then

        MsgBox "Erros encontrados:" & vbCrLf & vbCrLf & erros, vbExclamation

        GoTo Finalizar

    End If

    ' =========================
    ' FASE 2 - PROCESSAMENTO
    ' =========================
    For i = 2 To ultLinha

        cod = Trim(ws.Cells(i, 1).Value)

        If cod = "" Then GoTo ProxExecucao

        lote = Trim(ws.Cells(i, 4).Value)
        qty = CDbl(ws.Cells(i, 3).Value)

        ultEstoque = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

        For j = ultEstoque To 2 Step -1

            If wsEstoque.Cells(j, 1).Value = cod And _
               wsEstoque.Cells(j, 3).Value = lote Then

                Dim livreExec As Double
                Dim bloqExec As Double

                livreExec = Val(wsEstoque.Cells(j, 4).Value)
                bloqExec = Val(wsEstoque.Cells(j, 6).Value)

                wsEstoque.Cells(j, 4).Value = livreExec - qty
                wsEstoque.Cells(j, 6).Value = bloqExec + qty

                Dim h As Long

                h = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1

                wsHist.Cells(h, 1).Value = cod
                wsHist.Cells(h, 2).Value = wsEstoque.Cells(j, 2).Value
                wsHist.Cells(h, 3).Value = lote
                wsHist.Cells(h, 4).Value = qty
                wsHist.Cells(h, 5).Value = "Bloqueio"
                wsHist.Cells(h, 6).Value = Environ("USERNAME")
                wsHist.Cells(h, 7).Value = Now()
                wsHist.Cells(h, 8).Value = wsEstoque.Cells(j, 8).Value

                Call ColorirLinhaHistorico(wsHist, h, "Bloqueio")

                Exit For

            End If

        Next j

ProxExecucao:
    Next i

    ws.Range("A2:D" & ultLinha).ClearContents
    ws.Range("A2:D" & ultLinha).Interior.ColorIndex = xlNone

    Call OrdenarEstoque
    Call FiltrarEstoque
    Call OrdenarHistorico
    Call SalvarSeguro
    

    MsgBox "Bloqueios realizados com sucesso!", vbInformation

Finalizar:
    If Err.Number <> 0 Then

    MsgBox _
        "Erro " & Err.Number & vbCrLf & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Sub: BloquearMaterial", _
        vbCritical

    End If

    Call ProtegerSistema

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
' ==============================================
' LIBERAR
' ==============================================
Sub LiberarMaterial()

    On Error GoTo Finalizar

    If SENHA = "" Then SENHA = GetSenha()
    Call DesprotegerSistema

    Dim ws As Worksheet
    Dim wsEstoque As Worksheet
    Dim wsHist As Worksheet

    Set ws = ThisWorkbook.Sheets("Bloqueados")
    Set wsEstoque = ThisWorkbook.Sheets("Estoque Atual")
    Set wsHist = ThisWorkbook.Sheets("Historico")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim ultLinha As Long
    ultLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim temDados As Boolean
    Dim x As Long

    For x = 2 To ultLinha

        If Trim(ws.Cells(x, 1).Value) <> "" Then
            temDados = True
            Exit For
        End If

    Next x

    If Not temDados Then
        MsgBox "Nenhuma linha preenchida!", vbExclamation
        GoTo Finalizar
    End If

    Dim erros As String
    Dim temErro As Boolean
    Dim i As Long

    ' =========================
    ' LIMPA CORES ANTIGAS
    ' =========================
    For i = 2 To ultLinha
        ws.Range("A" & i & ":D" & i).Interior.ColorIndex = xlNone
    Next i

    ' =========================
    ' FASE 1 - VALIDAÇÃO
    ' =========================
    For i = 2 To ultLinha

        Dim cod As String
        Dim lote As String
        Dim qty As Double
        Dim encontrado As Boolean

        encontrado = False

        cod = Trim(ws.Cells(i, 1).Value)
        lote = Trim(ws.Cells(i, 4).Value)

        If cod = "" And lote = "" Then GoTo ProxValidacao

        If cod = "" Or lote = "" Then

            erros = erros & "Linha " & i & ": Código ou lote vazio" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValidacao

        End If

        If Not IsNumeric(ws.Cells(i, 3).Value) Then

            erros = erros & "Linha " & i & ": Quantidade inválida" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValidacao

        End If

        qty = CDbl(ws.Cells(i, 3).Value)

        If qty <= 0 Then

            erros = erros & "Linha " & i & ": Quantidade <= 0" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValidacao

        End If

        Dim j As Long
        Dim ult As Long

        ult = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

        For j = 2 To ult

            If wsEstoque.Cells(j, 1).Value = cod And _
               wsEstoque.Cells(j, 3).Value = lote Then

                encontrado = True

                Dim saldoBloq As Double
                saldoBloq = Val(wsEstoque.Cells(j, 6).Value)

                If qty > saldoBloq Then

                    erros = erros & "Linha " & i & ": Saldo bloqueado insuficiente" & vbCrLf

                    ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

                    temErro = True

                End If

                Exit For

            End If

        Next j

        If Not encontrado Then

            erros = erros & "Linha " & i & ": Lote não encontrado" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

        End If

ProxValidacao:
    Next i

    If temErro Then

        MsgBox "Erros encontrados:" & vbCrLf & vbCrLf & erros, vbExclamation

        GoTo Finalizar

    End If

    ' =========================
    ' FASE 2 - PROCESSAMENTO
    ' =========================
    For i = 2 To ultLinha

        cod = Trim(ws.Cells(i, 1).Value)

        If cod = "" Then GoTo ProxExecucao

        lote = Trim(ws.Cells(i, 4).Value)
        qty = CDbl(ws.Cells(i, 3).Value)

        ult = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

        For j = 2 To ult

            If wsEstoque.Cells(j, 1).Value = cod And _
               wsEstoque.Cells(j, 3).Value = lote Then

                Dim livre As Double
                Dim bloq As Double

                livre = Val(wsEstoque.Cells(j, 4).Value)
                bloq = Val(wsEstoque.Cells(j, 6).Value)

                wsEstoque.Cells(j, 4).Value = livre + qty
                wsEstoque.Cells(j, 6).Value = bloq - qty

                Dim h As Long

                h = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1

                wsHist.Cells(h, 1).Value = cod
                wsHist.Cells(h, 2).Value = wsEstoque.Cells(j, 2).Value
                wsHist.Cells(h, 3).Value = lote
                wsHist.Cells(h, 4).Value = qty
                wsHist.Cells(h, 5).Value = "Liberação"
                wsHist.Cells(h, 6).Value = Environ("USERNAME")
                wsHist.Cells(h, 7).Value = Now()
                wsHist.Cells(h, 8).Value = wsEstoque.Cells(j, 8).Value

                Call ColorirLinhaHistorico(wsHist, h, "Liberação")

                Exit For

            End If

        Next j

ProxExecucao:
    Next i

    ws.Range("A2:D" & ultLinha).ClearContents
    ws.Range("A2:D" & ultLinha).Interior.ColorIndex = xlNone

    Call OrdenarEstoque
    Call FiltrarEstoque
    Call OrdenarHistorico
    Call SalvarSeguro
    

    MsgBox "Liberação finalizada!", vbInformation

Finalizar:
    If Err.Number <> 0 Then

    MsgBox _
        "Erro " & Err.Number & vbCrLf & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Sub: LiberarMaterial", _
        vbCritical

    End If

    Call ProtegerSistema

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub FiltrarEstoque()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Estoque Atual")

    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim i As Long

    ' =========================
    ' CALCULA TOTAL (coluna G)
    ' =========================
    For i = 2 To ult

        Dim livre As Double
        Dim qualidade As Double
        Dim bloqueado As Double
        Dim total As Double

        livre = Val(ws.Cells(i, 4).Value)
        qualidade = Val(ws.Cells(i, 5).Value)
        bloqueado = Val(ws.Cells(i, 6).Value)

        total = livre + qualidade + bloqueado

        ws.Cells(i, 7).Value = total

    Next i

    ' =========================
    ' REMOVE FILTRO ANTIGO
    ' =========================
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    ' =========================
    ' APLICA FILTRO (TOTAL > 0)
    ' =========================
    ws.Range("A1:G" & ult).AutoFilter Field:=7, Criteria1:=">0"

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
Sub OrdenarEstoque()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Estoque Atual")

    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    On Error GoTo Fim
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Call DesprotegerSistema

    ' ?? REMOVE FILTRO ANTES DE ORDENAR
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("A2:A" & ult), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("C2:C" & ult), Order:=xlAscending
        .SetRange ws.Range("A1:H" & ult)
        .Header = xlYes
        .Apply
    End With

    Call ProtegerSistema

    ' limpa seleção visual
    ws.Activate
    ws.Range("A1").Select
    
Fim:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
Sub ConsultarMovimentacao()

    Dim wsHist As Worksheet
    Dim wsConsulta As Worksheet
    
    Set wsHist = ThisWorkbook.Sheets("Historico")
    Set wsConsulta = ThisWorkbook.Sheets("Consultar Movimentação")
    
    Dim ultHist As Long
    ultHist = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row
    
    ' limpa resultados antigos + cores
    wsConsulta.Range("A9:J1000").ClearContents
    wsConsulta.Range("A9:J1000").Interior.ColorIndex = xlNone
    
    Dim linhaDestino As Long
    linhaDestino = 9
    
    Dim i As Long
    
    For i = 2 To ultHist
    
        Dim cod As String
        Dim usuario As String
        Dim tipo As String
        Dim dataMov As Variant
        Dim notaFiscal As String
        Dim lote As String
        
        lote = Trim(wsHist.Cells(i, 3).Value)
        cod = wsHist.Cells(i, 1).Value
        usuario = wsHist.Cells(i, 6).Value
        tipo = wsHist.Cells(i, 5).Value
        dataMov = wsHist.Cells(i, 7).Value
        notaFiscal = wsHist.Cells(i, 8).Value
        
        ' filtros
        Dim dataFiltro As Variant
        dataFiltro = wsConsulta.Range("C2").Value

        If (wsConsulta.Range("A2").Value = "" Or cod = wsConsulta.Range("A2").Value) _
            And (wsConsulta.Range("B2").Value = "" Or usuario = wsConsulta.Range("B2").Value) _
            And (wsConsulta.Range("D2").Value = "" Or tipo = wsConsulta.Range("D2").Value) _
            And (wsConsulta.Range("E2").Value = "" Or notaFiscal = wsConsulta.Range("E2").Value) _
            And (wsConsulta.Range("F2").Value = "" Or Trim(wsHist.Cells(i, 9).Value) = Trim(wsConsulta.Range("F2").Value)) _
            And (wsConsulta.Range("G2").Value = "" Or Trim(wsHist.Cells(i, 10).Value) = Trim(wsConsulta.Range("G2").Value)) _
            And (wsConsulta.Range("H2").Value = "" Or lote = Trim(wsConsulta.Range("H2").Value)) _
            And ( _
        IsEmpty(dataFiltro) _
            Or (IsDate(dataFiltro) And DateValue(CDate(dataMov)) = DateValue(CDate(dataFiltro))) _
) Then
        
            ' copia linha
            wsConsulta.Cells(linhaDestino, 1).Value = wsHist.Cells(i, 1).Value
            wsConsulta.Cells(linhaDestino, 2).Value = wsHist.Cells(i, 2).Value
            wsConsulta.Cells(linhaDestino, 3).Value = wsHist.Cells(i, 3).Value
            wsConsulta.Cells(linhaDestino, 4).Value = wsHist.Cells(i, 4).Value
            wsConsulta.Cells(linhaDestino, 5).Value = wsHist.Cells(i, 5).Value
            wsConsulta.Cells(linhaDestino, 6).Value = wsHist.Cells(i, 6).Value
            wsConsulta.Cells(linhaDestino, 7).Value = wsHist.Cells(i, 7).Value
            wsConsulta.Cells(linhaDestino, 8).Value = wsHist.Cells(i, 8).Value
            wsConsulta.Cells(linhaDestino, 9).Value = wsHist.Cells(i, 10).Value ' Transf Para
            wsConsulta.Cells(linhaDestino, 10).Value = wsHist.Cells(i, 9).Value ' Pedido
            
            
            Select Case UCase(Trim(tipo))
                Case "ENTRADA"
                    wsConsulta.Range("A" & linhaDestino & ":J" & linhaDestino).Interior.Color = RGB(198, 239, 206) ' verde
                
                Case "BAIXA"
                    wsConsulta.Range("A" & linhaDestino & ":J" & linhaDestino).Interior.Color = RGB(255, 199, 206) ' vermelho
                
                Case "BLOQUEIO"
                    wsConsulta.Range("A" & linhaDestino & ":J" & linhaDestino).Interior.Color = RGB(255, 217, 102) ' laranja
                
                Case "LIBERAÇÃO"
                    wsConsulta.Range("A" & linhaDestino & ":J" & linhaDestino).Interior.Color = RGB(255, 255, 153) ' amarelo
                Case "TRANSFERÊNCIA"
                    wsConsulta.Range("A" & linhaDestino & ":J" & linhaDestino).Interior.Color = RGB(189, 215, 238) ' azul
                Case "CONTROLE QUALIDADE"
                    wsConsulta.Range("A" & linhaDestino & ":J" & linhaDestino).Interior.Color = RGB(255, 230, 153) ' amarelo suave
                Case "LIBERAÇÃO CQ", "LIBERACAO CQ"
                    wsConsulta.Range("A" & linhaDestino & ":J" & linhaDestino).Interior.Color = RGB(255, 255, 153) ' mesmo do histórico
                Case "ESTORNO BAIXA"
                    wsConsulta.Range("A" & linhaDestino & ":J" & linhaDestino).Interior.Color = RGB(217, 225, 242)
            End Select
            
            linhaDestino = linhaDestino + 1
            
        End If
        
    Next i
    
    wsConsulta.Range("A9").Select
    MsgBox "Consulta finalizada!", vbInformation

End Sub


Sub ColorirLinhaHistorico(ws As Worksheet, linha As Long, tipo As String)

    With ws.Range("A" & linha & ":J" & linha)
    
        Select Case UCase(tipo)
        
            Case "ENTRADA"
                .Interior.Color = RGB(198, 239, 206) ' verde
                
            Case "BAIXA"
                .Interior.Color = RGB(255, 199, 206) ' vermelho
                
            Case "BLOQUEIO"
                .Interior.Color = RGB(255, 217, 102) ' laranja
                
            Case "LIBERAÇÃO"
                .Interior.Color = RGB(255, 255, 153) ' amarelo
            Case "TRANSFERÊNCIA"
                .Interior.Color = RGB(180, 198, 231) '  azul leve
            Case "CONTROLE QUALIDADE"
                .Interior.Color = RGB(255, 230, 153) ' amarelo suave
            Case "LIBERAÇÃO CQ"
                .Interior.Color = RGB(189, 215, 238) ' azul clarinho
            Case "ESTORNO BAIXA"
                .Interior.Color = RGB(217, 225, 242)
                
        End Select
        
    End With

End Sub


Sub RecolorirHistorico()

    On Error GoTo Finalizar

    If SENHA = "" Then SENHA = GetSenha()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Historico")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' =========================
    ' DESPROTEGE
    ' =========================
    ws.Unprotect Password:=SENHA

    Dim ultLinha As Long
    ultLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    Dim tipo As String

    ' =========================
    ' LIMPA CORES
    ' =========================
    ws.Range("A2:J1000").Interior.ColorIndex = xlNone

    ' =========================
    ' RECOLORIR
    ' =========================
    For i = 2 To ultLinha

        If Trim(ws.Cells(i, 1).Value) <> "" Then

            tipo = Trim(ws.Cells(i, 5).Value)

            Call ColorirLinhaHistorico(ws, i, tipo)

        End If

    Next i

Finalizar:

    ' =========================
    ' REPROTEGE
    ' =========================
    ws.Protect Password:=SENHA, UserInterfaceOnly:=True

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


Sub RegistrarTransferencia()

    On Error GoTo Finalizar

    If SENHA = "" Then SENHA = GetSenha()
    Call DesprotegerSistema

    Dim ws As Worksheet
    Dim wsEstoque As Worksheet
    Dim wsHist As Worksheet

    Set ws = ThisWorkbook.Sheets("Transferência")
    Set wsEstoque = ThisWorkbook.Sheets("Estoque Atual")
    Set wsHist = ThisWorkbook.Sheets("Historico")

    Dim i As Long
    Dim j As Long
    Dim ult As Long

    Dim temDados As Boolean
    Dim temErro As Boolean
    Dim erros As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' =========================
    ' VALIDA SE TEM DADOS
    ' =========================
    For i = 6 To 60

        If Trim(ws.Cells(i, 2).Value) <> "" Then
            temDados = True
            Exit For
        End If

    Next i

    If Not temDados Then
        MsgBox "Nenhum item para transferir!", vbExclamation
        GoTo Finalizar
    End If

    ' =========================
    ' DESTINO / PEDIDO
    ' =========================
    Dim destino As String
    Dim pedido As String

    destino = Trim(ws.Range("C2").Value)
    pedido = Trim(ws.Range("E2").Value)

    If destino = "" Or pedido = "" Then
        MsgBox "Preencha destino e pedido!", vbExclamation
        GoTo Finalizar
    End If

    ult = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

    ' =========================
    ' LIMPA CORES ANTIGAS
    ' =========================
    For i = 6 To 60
        ws.Range("B" & i & ":E" & i).Interior.ColorIndex = xlNone
    Next i

    ' =========================
    ' FASE 1 - VALIDAÇÃO
    ' =========================
    For i = 6 To 60

        Dim codExec As String
        Dim loteExec As String
        Dim qtyExec As Double
        Dim encontrado As Boolean

        encontrado = False

        codExec = Trim(ws.Cells(i, 2).Value)
        loteExec = Trim(ws.Cells(i, 5).Value)

        If codExec = "" And loteExec = "" Then GoTo ProxValidacao

        ' código/lote vazio
        If codExec = "" Or loteExec = "" Then

            erros = erros & "Linha " & i & ": Código ou lote vazio" & vbCrLf

            ws.Range("B" & i & ":E" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValidacao

        End If

        ' quantidade inválida
        If Not IsNumeric(ws.Cells(i, 4).Value) Then

            erros = erros & "Linha " & i & ": Quantidade inválida" & vbCrLf

            ws.Range("B" & i & ":E" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValidacao

        End If

        qtyExec = CDbl(ws.Cells(i, 4).Value)

        ' quantidade <= 0
        If qtyExec <= 0 Then

            erros = erros & "Linha " & i & ": Quantidade <= 0" & vbCrLf

            ws.Range("B" & i & ":E" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValidacao

        End If

        ' =========================
        ' VALIDA LOTE
        ' =========================
        For j = ult To 2 Step -1

            If wsEstoque.Cells(j, 1).Value = codExec And _
               wsEstoque.Cells(j, 3).Value = loteExec Then

                encontrado = True

                Dim saldoLivre As Double
                saldoLivre = Val(wsEstoque.Cells(j, 4).Value)

                If qtyExec > saldoLivre Then

                    erros = erros & "Linha " & i & ": Saldo insuficiente" & vbCrLf

                    ws.Range("B" & i & ":E" & i).Interior.Color = RGB(255, 200, 200)

                    temErro = True

                End If

                Exit For

            End If

        Next j

        ' lote não encontrado
        If Not encontrado Then

            erros = erros & "Linha " & i & ": Lote não encontrado" & vbCrLf

            ws.Range("B" & i & ":E" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

        End If

ProxValidacao:
    Next i

    ' =========================
    ' SE TIVER ERRO -> PARA TUDO
    ' =========================
    If temErro Then

        MsgBox "Erros encontrados:" & vbCrLf & vbCrLf & erros, vbExclamation

        GoTo Finalizar

    End If

    ' =========================
    ' FASE 2 - PROCESSAMENTO
    ' =========================
    For i = 6 To 60

        codExec = Trim(ws.Cells(i, 2).Value)

        If codExec = "" Then GoTo ProxExecucao

        loteExec = Trim(ws.Cells(i, 5).Value)
        qtyExec = CDbl(ws.Cells(i, 4).Value)

        ult = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

        For j = ult To 2 Step -1

            If wsEstoque.Cells(j, 1).Value = codExec And _
               wsEstoque.Cells(j, 3).Value = loteExec Then

                ' baixa estoque
                wsEstoque.Cells(j, 4).Value = wsEstoque.Cells(j, 4).Value - qtyExec

                Dim livre As Double
                Dim qualidade As Double
                Dim bloqueado As Double

                livre = Val(wsEstoque.Cells(j, 4).Value)
                qualidade = Val(wsEstoque.Cells(j, 5).Value)
                bloqueado = Val(wsEstoque.Cells(j, 6).Value)

                ' histórico
                Dim h As Long

                h = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1

                wsHist.Cells(h, 1).Value = codExec
                wsHist.Cells(h, 2).Value = wsEstoque.Cells(j, 2).Value
                wsHist.Cells(h, 3).Value = loteExec
                wsHist.Cells(h, 4).Value = qtyExec
                wsHist.Cells(h, 5).Value = "Transferência"
                wsHist.Cells(h, 6).Value = Environ("USERNAME")
                wsHist.Cells(h, 7).Value = Now()
                wsHist.Cells(h, 9).Value = pedido
                wsHist.Cells(h, 10).Value = destino

                Call ColorirLinhaHistorico(wsHist, h, "Transferência")

                ' remove linha zerada
                If livre + qualidade + bloqueado = 0 Then
                    wsEstoque.Rows(j).Delete
                End If

                Exit For

            End If

        Next j

ProxExecucao:
    Next i

    ' =========================
    ' LIMPA SOMENTE DADOS
    ' =========================
    ws.Range("B6:E60").ClearContents
    ws.Range("C2").ClearContents
    ws.Range("E2").ClearContents
    
    Call OrdenarEstoque
    Call FiltrarEstoque
    Call OrdenarHistorico
    Call SalvarSeguro
    

    MsgBox "Transferência realizada com sucesso!", vbInformation

Finalizar:
    If Err.Number <> 0 Then

    MsgBox _
        "Erro " & Err.Number & vbCrLf & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Sub: RegistrarTransferencia", _
        vbCritical

    End If

    Call ProtegerSistema

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
Sub RegistrarControleQualidade()

    On Error GoTo Finalizar

    If SENHA = "" Then SENHA = GetSenha()
    Call DesprotegerSistema

    Dim ws As Worksheet
    Dim wsEstoque As Worksheet
    Dim wsHist As Worksheet

    Set ws = ThisWorkbook.Sheets("Controle de Qualidade")
    Set wsEstoque = ThisWorkbook.Sheets("Estoque Atual")
    Set wsHist = ThisWorkbook.Sheets("Historico")

    Dim erros As String
    Dim temErro As Boolean
    Dim i As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim temDados As Boolean

    For i = 2 To 50

        If Trim(ws.Cells(i, 1).Value) <> "" Then
            temDados = True
            Exit For
        End If

    Next i

    If Not temDados Then
        MsgBox "Nenhum item informado!", vbExclamation
        GoTo Finalizar
    End If

    ' =========================
    ' VALIDAÇÃO
    ' =========================
    For i = 2 To 50

        Dim cod As String
        Dim lote As String
        Dim qty As Double

        cod = Trim(ws.Cells(i, 1).Value)
        lote = Trim(ws.Cells(i, 4).Value)

        If cod = "" And lote = "" Then GoTo ProxVal

        ws.Range("A" & i & ":D" & i).Interior.ColorIndex = xlNone

        If cod = "" Or lote = "" Then

            erros = erros & "Linha " & i & ": Código ou lote vazio" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxVal

        End If

        If Not IsNumeric(ws.Cells(i, 3).Value) Then

            erros = erros & "Linha " & i & ": Quantidade inválida" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxVal

        End If

        qty = CDbl(ws.Cells(i, 3).Value)

        If qty <= 0 Then

            erros = erros & "Linha " & i & ": Quantidade <= 0" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxVal

        End If

        Dim j As Long
        Dim ult As Long

        ult = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

        Dim encontrado As Boolean

        encontrado = False

        For j = 2 To ult

            If wsEstoque.Cells(j, 1).Value = cod And _
               wsEstoque.Cells(j, 3).Value = lote Then

                encontrado = True

                If qty > Val(wsEstoque.Cells(j, 4).Value) Then

                    erros = erros & "Linha " & i & ": Saldo livre insuficiente" & vbCrLf

                    ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

                    temErro = True

                End If

                Exit For

            End If

        Next j

        If Not encontrado Then

            erros = erros & "Linha " & i & ": Lote não encontrado" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

        End If

ProxVal:
    Next i

    If temErro Then

        MsgBox "Erros encontrados:" & vbCrLf & vbCrLf & erros, vbExclamation

        GoTo Finalizar

    End If

    ' =========================
    ' PROCESSAMENTO
    ' =========================
    For i = 2 To 50

        cod = Trim(ws.Cells(i, 1).Value)
        lote = Trim(ws.Cells(i, 4).Value)

        If cod = "" Then GoTo ProxExec

        qty = CDbl(ws.Cells(i, 3).Value)

        For j = 2 To wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

            If wsEstoque.Cells(j, 1).Value = cod And _
               wsEstoque.Cells(j, 3).Value = lote Then

                wsEstoque.Cells(j, 4).Value = Val(wsEstoque.Cells(j, 4).Value) - qty
                wsEstoque.Cells(j, 5).Value = Val(wsEstoque.Cells(j, 5).Value) + qty

                Dim livre As Double
                Dim qualidade As Double
                Dim bloqueado As Double

                livre = Val(wsEstoque.Cells(j, 4).Value)
                qualidade = Val(wsEstoque.Cells(j, 5).Value)
                bloqueado = Val(wsEstoque.Cells(j, 6).Value)

                Dim h As Long

                h = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1

                wsHist.Cells(h, 1).Value = cod
                wsHist.Cells(h, 2).Value = wsEstoque.Cells(j, 2).Value
                wsHist.Cells(h, 3).Value = lote
                wsHist.Cells(h, 4).Value = qty
                wsHist.Cells(h, 5).Value = "Controle Qualidade"
                wsHist.Cells(h, 6).Value = Environ("USERNAME")
                wsHist.Cells(h, 7).Value = Now()

                Call ColorirLinhaHistorico(wsHist, h, "Controle Qualidade")

                If livre + qualidade + bloqueado = 0 Then
                    wsEstoque.Rows(j).Delete
                End If

                Exit For

            End If

        Next j

ProxExec:
    Next i

    ws.Range("A2:D50").ClearContents
    ws.Range("A2:D50").Interior.ColorIndex = xlNone

    Call OrdenarEstoque
    Call FiltrarEstoque
    Call OrdenarHistorico
    Call SalvarSeguro
    

    MsgBox "Movimentação para Controle de Qualidade realizada!", vbInformation

Finalizar:
    If Err.Number <> 0 Then

    MsgBox _
        "Erro " & Err.Number & vbCrLf & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Sub: RegistrarControleQualidade", _
        vbCritical

    End If

    Call ProtegerSistema

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
Sub LiberarControleQualidade()

    On Error GoTo Finalizar

    If SENHA = "" Then SENHA = GetSenha()
    Call DesprotegerSistema

    Dim ws As Worksheet
    Dim wsEstoque As Worksheet
    Dim wsHist As Worksheet

    Set ws = ThisWorkbook.Sheets("Controle de Qualidade")
    Set wsEstoque = ThisWorkbook.Sheets("Estoque Atual")
    Set wsHist = ThisWorkbook.Sheets("Historico")

    Dim erros As String
    Dim temErro As Boolean
    Dim temDados As Boolean
    Dim i As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = 2 To 50

        If Trim(ws.Cells(i, 1).Value) <> "" Then
            temDados = True
            Exit For
        End If

    Next i

    If Not temDados Then
        MsgBox "Nenhum item informado!", vbExclamation
        GoTo Finalizar
    End If

    ' =========================
    ' VALIDAÇÃO
    ' =========================
    For i = 2 To 50

        Dim cod As String
        Dim lote As String
        Dim qty As Double

        cod = Trim(ws.Cells(i, 1).Value)
        lote = Trim(ws.Cells(i, 4).Value)

        If cod = "" And lote = "" Then GoTo ProxVal

        ws.Range("A" & i & ":D" & i).Interior.ColorIndex = xlNone

        If cod = "" Or lote = "" Then

            erros = erros & "Linha " & i & ": Código ou lote vazio" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxVal

        End If

        If Not IsNumeric(ws.Cells(i, 3).Value) Then

            erros = erros & "Linha " & i & ": Quantidade inválida" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxVal

        End If

        qty = CDbl(ws.Cells(i, 3).Value)

        If qty <= 0 Then

            erros = erros & "Linha " & i & ": Quantidade <= 0" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxVal

        End If

        Dim j As Long
        Dim ult As Long

        ult = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

        Dim encontrado As Boolean

        encontrado = False

        For j = 2 To ult

            If wsEstoque.Cells(j, 1).Value = cod And _
               wsEstoque.Cells(j, 3).Value = lote Then

                encontrado = True

                If qty > Val(wsEstoque.Cells(j, 5).Value) Then

                    erros = erros & "Linha " & i & ": Saldo em CQ insuficiente" & vbCrLf

                    ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

                    temErro = True

                End If

                Exit For

            End If

        Next j

        If Not encontrado Then

            erros = erros & "Linha " & i & ": Lote não encontrado" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

        End If

ProxVal:
    Next i

    If temErro Then

        MsgBox "Erros encontrados:" & vbCrLf & vbCrLf & erros, vbExclamation

        GoTo Finalizar

    End If

    ' =========================
    ' PROCESSAMENTO
    ' =========================
    For i = 2 To 50

        cod = Trim(ws.Cells(i, 1).Value)
        lote = Trim(ws.Cells(i, 4).Value)

        If cod = "" Then GoTo ProxExec

        qty = CDbl(ws.Cells(i, 3).Value)

        For j = 2 To wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

            If wsEstoque.Cells(j, 1).Value = cod And _
               wsEstoque.Cells(j, 3).Value = lote Then

                wsEstoque.Cells(j, 5).Value = Val(wsEstoque.Cells(j, 5).Value) - qty
                wsEstoque.Cells(j, 4).Value = Val(wsEstoque.Cells(j, 4).Value) + qty

                Dim livre As Double
                Dim qualidade As Double
                Dim bloqueado As Double

                livre = Val(wsEstoque.Cells(j, 4).Value)
                qualidade = Val(wsEstoque.Cells(j, 5).Value)
                bloqueado = Val(wsEstoque.Cells(j, 6).Value)

                Dim h As Long

                h = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1

                wsHist.Cells(h, 1).Value = cod
                wsHist.Cells(h, 2).Value = wsEstoque.Cells(j, 2).Value
                wsHist.Cells(h, 3).Value = lote
                wsHist.Cells(h, 4).Value = qty
                wsHist.Cells(h, 5).Value = "Liberação CQ"
                wsHist.Cells(h, 6).Value = Environ("USERNAME")
                wsHist.Cells(h, 7).Value = Now()

                Call ColorirLinhaHistorico(wsHist, h, "Liberação CQ")

                If livre + qualidade + bloqueado = 0 Then
                    wsEstoque.Rows(j).Delete
                End If

                Exit For

            End If

        Next j

ProxExec:
    Next i

    ws.Range("A2:D50").ClearContents
    ws.Range("A2:D50").Interior.ColorIndex = xlNone

    Call OrdenarEstoque
    Call FiltrarEstoque
    Call OrdenarHistorico
    Call SalvarSeguro
    

    MsgBox "Liberação do Controle de Qualidade realizada!", vbInformation

Finalizar:
    If Err.Number <> 0 Then

    MsgBox _
        "Erro " & Err.Number & vbCrLf & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Sub: LiberarControleQualidade", _
        vbCritical

    End If

    Call ProtegerSistema

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
Function SaldoDisponivelEstorno(cod As String, lote As String) As Double

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Historico")

    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim i As Long
    Dim totalSaida As Double
    Dim totalEstorno As Double

    For i = 2 To ult

        If ws.Cells(i, 1).Value = cod And ws.Cells(i, 3).Value = lote Then

            Select Case UCase(Trim(ws.Cells(i, 5).Value))

                Case "BAIXA", "TRANSFERÊNCIA", "TRANSFERENCIA"
                    totalSaida = totalSaida + ws.Cells(i, 4).Value

                Case "ESTORNO BAIXA"
                    totalEstorno = totalEstorno + ws.Cells(i, 4).Value

            End Select

        End If

    Next i

    SaldoDisponivelEstorno = totalSaida - totalEstorno

End Function
Sub EstornarBaixa()

    On Error GoTo Finalizar

    If SENHA = "" Then SENHA = GetSenha()

    Call DesprotegerSistema

    Dim ws As Worksheet
    Dim wsEstoque As Worksheet
    Dim wsHist As Worksheet

    Set ws = ThisWorkbook.Sheets("Estornar Baixa")
    Set wsEstoque = ThisWorkbook.Sheets("Estoque Atual")
    Set wsHist = ThisWorkbook.Sheets("Historico")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim ultLinha As Long
    ultLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim temDados As Boolean
    Dim i As Long

    ' =========================
    ' VERIFICA DADOS
    ' =========================
    For i = 2 To ultLinha

        If Trim(ws.Cells(i, 1).Value) <> "" Then
            temDados = True
            Exit For
        End If

    Next i

    If Not temDados Then
        MsgBox "Nenhuma linha preenchida!", vbExclamation
        GoTo Finalizar
    End If

    Dim erros As String
    Dim temErro As Boolean

    ' =========================
    ' VALIDAÇÃO
    ' =========================
    For i = 2 To ultLinha

        Dim cod As String
        Dim lote As String
        Dim qtd As Double

        cod = Trim(ws.Cells(i, 1).Value)
        lote = Trim(ws.Cells(i, 4).Value)

        ws.Range("A" & i & ":D" & i).Interior.ColorIndex = xlNone

        If cod = "" And lote = "" Then GoTo ProxValid

        If cod = "" Or lote = "" Then

            erros = erros & "Linha " & i & ": Código ou lote vazio" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValid

        End If

        If Not IsNumeric(ws.Cells(i, 3).Value) Then

            erros = erros & "Linha " & i & ": Quantidade inválida" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValid

        End If

        qtd = CDbl(ws.Cells(i, 3).Value)

        If qtd <= 0 Then

            erros = erros & "Linha " & i & ": Quantidade <= 0" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

            GoTo ProxValid

        End If

        Dim disponivel As Double

        disponivel = SaldoDisponivelEstorno(cod, lote)

        If qtd > disponivel Then

            erros = erros & "Linha " & i & _
            ": Estorno maior que permitido (Disp: " & disponivel & ")" & vbCrLf

            ws.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 200, 200)

            temErro = True

        End If

ProxValid:
    Next i

    If temErro Then

        MsgBox "Erros encontrados:" & vbCrLf & vbCrLf & erros, vbExclamation

        GoTo Finalizar

    End If

    ' =========================
    ' PROCESSAMENTO
    ' =========================
    For i = 2 To ultLinha

        cod = Trim(ws.Cells(i, 1).Value)

        If cod = "" Then GoTo ProxExec

        lote = Trim(ws.Cells(i, 4).Value)
        qtd = CDbl(ws.Cells(i, 3).Value)

        Dim j As Long
        Dim ultEst As Long
        Dim encontrado As Boolean

        encontrado = False

        ultEst = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row

        For j = 2 To ultEst

            If wsEstoque.Cells(j, 1).Value = cod _
            And wsEstoque.Cells(j, 3).Value = lote Then

                encontrado = True

                wsEstoque.Cells(j, 4).Value = wsEstoque.Cells(j, 4).Value + qtd

                Exit For

            End If

        Next j

        ' =========================
        ' RECRIA LOTE SE NÃO EXISTIR
        ' =========================
        If Not encontrado Then

            Dim novaLinha As Long

            novaLinha = wsEstoque.Cells(wsEstoque.Rows.Count, "A").End(xlUp).Row + 1

            Dim descMaterial As String
            descMaterial = BuscarDescricaoRapida(cod)

            Dim notaFiscal As String
            notaFiscal = ""

            ' tenta pegar NF do histórico
            Dim hBusca As Long

            For hBusca = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row To 2 Step -1

                If wsHist.Cells(hBusca, 1).Value = cod _
                And wsHist.Cells(hBusca, 3).Value = lote Then

                    notaFiscal = wsHist.Cells(hBusca, 8).Value

                    Exit For

                End If

            Next hBusca

            wsEstoque.Cells(novaLinha, 1).Value = cod
            wsEstoque.Cells(novaLinha, 2).Value = descMaterial
            wsEstoque.Cells(novaLinha, 3).Value = lote
            wsEstoque.Cells(novaLinha, 4).Value = qtd
            wsEstoque.Cells(novaLinha, 5).Value = 0
            wsEstoque.Cells(novaLinha, 6).Value = 0
            wsEstoque.Cells(novaLinha, 8).Value = notaFiscal

        End If

        ' =========================
        ' HISTÓRICO
        ' =========================
        Dim h As Long

        h = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1

        wsHist.Cells(h, 1).Value = cod
        wsHist.Cells(h, 2).Value = BuscarDescricaoRapida(cod)
        wsHist.Cells(h, 3).Value = lote
        wsHist.Cells(h, 4).Value = qtd
        wsHist.Cells(h, 5).Value = "Estorno Baixa"
        wsHist.Cells(h, 6).Value = Environ("USERNAME")
        wsHist.Cells(h, 7).Value = Now()

        Call ColorirLinhaHistorico(wsHist, h, "Estorno Baixa")

ProxExec:
    Next i

    ws.Range("A2:D" & ultLinha).ClearContents
    ws.Range("A2:D" & ultLinha).Interior.ColorIndex = xlNone

    Call OrdenarEstoque
    Call FiltrarEstoque
    Call OrdenarHistorico
    Call SalvarSeguro
    
    MsgBox "Estorno realizado com sucesso!", vbInformation

Finalizar:

    If Err.Number <> 0 Then

        MsgBox _
            "Erro " & Err.Number & vbCrLf & vbCrLf & _
            Err.Description & vbCrLf & vbCrLf & _
            "Sub: EstornarBaixa", _
            vbCritical

    End If

    Call ProtegerSistema

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub AbrirEstoqueAtual()
    Call AbrirAba("Estoque Atual")
End Sub

Sub AbrirEntradas()
    Call AbrirAba("Entradas")
End Sub

Sub AbrirBaixas()
    Call AbrirAba("Baixas")
End Sub

Sub AbrirTransferencia()
    Call AbrirAba("Transferência")
End Sub

Sub AbrirEstornarBaixa()
    Call AbrirAba("Estornar Baixa")
End Sub

Sub AbrirControleQualidade()
    Call AbrirAba("Controle de Qualidade")
End Sub

Sub AbrirBloqueados()
    Call AbrirAba("Bloqueados")
End Sub

Sub AbrirBuscaMaterial()
    Call AbrirAba("Busca Materiais")
End Sub

Sub AbrirConsultarMovimentacao()
    Call AbrirAba("Consultar Movimentação")
End Sub

Sub AbrirHistorico()
    Call AbrirAba("Historico")
End Sub

Sub AbrirMateriais()
    Call AbrirAba("Materiais")
End Sub

Sub AbrirExportacoes()
    Call AbrirAba("Exportações")
End Sub

Sub VoltarDashboard()

    Application.ScreenUpdating = False

    Dim wsAtual As Worksheet
    Set wsAtual = ActiveSheet

    ' ?? garante senha atual
    If SENHA = "" Then SENHA = GetSenha

    ' ?? tenta liberar estrutura
    If ThisWorkbook.ProtectStructure Then
        
        On Error Resume Next
        ThisWorkbook.Unprotect Password:=SENHA
        On Error GoTo 0
        
    End If

    ' ?? se ainda estiver protegido, aborta
    If ThisWorkbook.ProtectStructure Then
        MsgBox "Erro ao liberar estrutura do sistema.", vbCritical
        GoTo Fim
    End If

    ' vai pro dashboard
    With Sheets("Dashboard")
        .Visible = xlSheetVisible
        .Activate
    End With

    ' esconde aba atual
    If wsAtual.Name <> "Dashboard" Then
        wsAtual.Visible = xlVeryHidden
    End If

    ' ?? protege novamente
    ThisWorkbook.Protect Password:=SENHA, Structure:=True

Fim:
    Application.ScreenUpdating = True

End Sub

Sub ProtegerSistemaCompleto()

    Dim ws As Worksheet

    ' ?? pega senha atual
    SENHA = GetSenha()

    Call OtimizarON

    ' ?? garante que tudo está desbloqueado antes
    On Error Resume Next
    
    ThisWorkbook.Unprotect SENHA
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect SENHA
    Next ws

    On Error GoTo 0

    ' ?? PROTEGE TODAS AS PLANILHAS
    For Each ws In ThisWorkbook.Worksheets

        ws.Protect _
            Password:=SENHA, _
            UserInterfaceOnly:=True, _
            AllowFiltering:=True, _
            AllowSorting:=True

    Next ws

    ' ?? PROTEGE ESTRUTURA DO ARQUIVO
    ThisWorkbook.Protect Password:=SENHA, Structure:=True

    Call OtimizarOFF

    MsgBox "Sistema protegido com sucesso!", vbInformation

End Sub

Sub DesprotegerTodasPlanilhas()

    Dim ws As Worksheet
    Dim frm As New frmSenha
    Dim senhaDigitada As String

    ' ?? pega senha atual do sistema
    SENHA = GetSenha()

    ' ?? abre form (senha oculta)
    frm.Show

    If frm.Cancelado Then Exit Sub

    senhaDigitada = frm.senhaDigitada

    ' ? valida senha
    If senhaDigitada <> SENHA Then
        MsgBox "Senha incorreta! Acesso negado.", vbCritical
        Exit Sub
    End If

    Call OtimizarON

    Call DesprotegerSistema

    Call OtimizarOFF

    MsgBox "Todas as planilhas foram desprotegidas!", vbInformation

End Sub


Public Sub SalvarSeguro()

    On Error GoTo ErroSalvar

    If ThisWorkbook.Path <> "" Then
        ThisWorkbook.Save
    End If

    Exit Sub

ErroSalvar:

    MsgBox "Erro ao salvar arquivo: " & Err.Description, vbCritical

End Sub



Sub AutoSaveLoop()

    On Error Resume Next

    ' cancela qualquer agendamento anterior (evita duplicidade)
    Application.OnTime ProximoAutoSave, "ExecutarAutoSave", , False

    On Error GoTo 0

    ProximoAutoSave = Now + TimeValue("00:02:00") ' a cada 2 minutos

    Application.OnTime ProximoAutoSave, "ExecutarAutoSave"

End Sub

Sub ExecutarAutoSave()

    Call SalvarSeguro
    Call AutoSaveLoop

End Sub

Sub BackupAutomatico()

    On Error Resume Next

    Dim caminho As String
    caminho = ThisWorkbook.Path & "\Backup\"

    If Dir(caminho, vbDirectory) = "" Then MkDir caminho

    Dim nome As String
    nome = "Backup_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"

    ThisWorkbook.SaveCopyAs caminho & nome

End Sub

Sub GerarPDFEstoque_TOP()

    Dim wsOrigem As Worksheet
    Dim wsPDF As Worksheet
    
    Set wsOrigem = ThisWorkbook.Sheets("Estoque Atual")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    Sheets("TEMP_PDF1").Delete
    On Error GoTo 0
    
    ' ?? libera estrutura
Call DesprotegerSistema

Set wsPDF = ThisWorkbook.Sheets.Add
wsPDF.Name = "TEMP_PDF1"

' ?? trava novamente
Call ProtegerSistema
    
    ' =========================
    ' CABEÇALHO (ALINHADO PERFEITO ??)
    ' =========================

    ' NOVA no centro
    wsPDF.Range("A1:F1").Merge
    wsPDF.Range("A1").Value = "NOVA " & Format(Now, "hh:mm")
    
    With wsPDF.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    
    ' QUEBRA O MERGE SÓ PARA POSICIONAR A DATA
    wsPDF.Range("A1:F1").UnMerge
    
    ' DATA NA MESMA LINHA (ALINHADA DIREITA)
    wsPDF.Range("F1").Value = Format(Date, "dd/mm/yyyy")
    
    With wsPDF.Range("F1")
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlRight
    End With
    
    ' RECRIA O CENTRO DO NOVA SEM ATRAPALHAR A DATA
    wsPDF.Range("A1:E1").Merge
    
    With wsPDF.Range("A1:E1")
        .HorizontalAlignment = xlCenter
    End With
    
    ' =========================
    ' TÍTULOS
    ' =========================
    
    wsPDF.Range("A3").Value = "Material"
    wsPDF.Range("B3").Value = "Texto breve material"
    wsPDF.Range("C3").Value = "Lote"
    wsPDF.Range("D3").Value = "Utilização livre"
    wsPDF.Range("E3").Value = "Em contr.qualidade"
    wsPDF.Range("F3").Value = "Bloqueado"
    
    With wsPDF.Range("A3:F3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' =========================
    ' DADOS
    ' =========================
    
    Dim ultLinha As Long
    ultLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long, linhaDestino As Long
    linhaDestino = 4
    
    For i = 2 To ultLinha
    
        Dim livre As Double, cq As Double, bloq As Double
        
        livre = wsOrigem.Cells(i, 4).Value
        cq = wsOrigem.Cells(i, 5).Value
        bloq = wsOrigem.Cells(i, 6).Value
        
        If livre <> 0 Or cq <> 0 Or bloq <> 0 Then
            
            wsPDF.Cells(linhaDestino, 1).Value = wsOrigem.Cells(i, 1).Value
            wsPDF.Cells(linhaDestino, 2).Value = wsOrigem.Cells(i, 2).Value
            wsPDF.Cells(linhaDestino, 3).Value = Format(wsOrigem.Cells(i, 3).Value, "0000000000")
            
            wsPDF.Cells(linhaDestino, 4).Value = livre
            wsPDF.Cells(linhaDestino, 5).Value = cq
            wsPDF.Cells(linhaDestino, 6).Value = bloq
            
            linhaDestino = linhaDestino + 1
            
        End If
        
    Next i
    
    ' =========================
    ' FORMATAÇÃO
    ' =========================
    
    wsPDF.Range("A4:A" & linhaDestino - 1).HorizontalAlignment = xlCenter
    wsPDF.Range("B4:B" & linhaDestino - 1).HorizontalAlignment = xlLeft
    wsPDF.Range("C4:F" & linhaDestino - 1).HorizontalAlignment = xlCenter
    
    wsPDF.Range("D4:F" & linhaDestino - 1).NumberFormat = "#,##0.000"
    
    wsPDF.Columns("A").ColumnWidth = 15
    wsPDF.Columns("B").ColumnWidth = 40
    wsPDF.Columns("C").ColumnWidth = 15
    wsPDF.Columns("D").ColumnWidth = 18
    wsPDF.Columns("E").ColumnWidth = 22
    wsPDF.Columns("F").ColumnWidth = 15
    
    With wsPDF.Range("A3:F" & linhaDestino - 1).Borders
        .LineStyle = xlContinuous
    End With
    
    ' =========================
    ' CONFIGURAÇÃO PDF
    ' =========================
    
    With wsPDF.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .CenterVertically = True
        
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
    End With
    
    wsPDF.PageSetup.PrintArea = wsPDF.Range("A1:F" & linhaDestino - 1).Address
    
    ' =========================
    ' EXPORTA PDF
    ' =========================
    
    Dim caminho As String
    caminho = ThisWorkbook.Path & "\Posicao_Estoque_" & Format(Now, "yyyymmdd_hhmm") & ".pdf"
    
    wsPDF.ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminho
    
    ' ?? ABRE AUTOMATICAMENTE O PDF
    ThisWorkbook.FollowHyperlink caminho
    
    ' ?? libera
Call DesprotegerSistema

wsPDF.Delete

' ?? trava de novo
Call ProtegerSistema
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "POSIÇÃO DE ESTOQUE GERADA!", vbInformation

End Sub

Sub ExportarEstoque_Excel()

    Dim wsOrigem As Worksheet
    Dim wbNovo As Workbook
    Dim wsNovo As Worksheet
    
    Set wsOrigem = ThisWorkbook.Sheets("Estoque Atual")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' =========================
    ' CRIA NOVO ARQUIVO
    ' =========================
    
    Set wbNovo = Workbooks.Add
    Set wsNovo = wbNovo.Sheets(1)
    wsNovo.Name = "Posição de Estoque"
    
    ' =========================
    ' CABEÇALHO (IGUAL PDF ??)
    ' =========================
    
    wsNovo.Range("A1:E1").Merge
    wsNovo.Range("A1").Value = "NOVA " & Format(Now, "hh:mm")
    
    With wsNovo.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    
    wsNovo.Range("F1").Value = Format(Date, "dd/mm/yyyy")
    
    With wsNovo.Range("F1")
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlRight
    End With
    
    ' =========================
    ' TÍTULOS
    ' =========================
    
    wsNovo.Range("A3").Value = "Material"
    wsNovo.Range("B3").Value = "Texto breve material"
    wsNovo.Range("C3").Value = "Lote"
    wsNovo.Range("D3").Value = "Utilização livre"
    wsNovo.Range("E3").Value = "Em contr.qualidade"
    wsNovo.Range("F3").Value = "Bloqueado"
    
    With wsNovo.Range("A3:F3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 220, 220)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' =========================
    ' DADOS
    ' =========================
    
    Dim ultLinha As Long
    ultLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long, linhaDestino As Long
    linhaDestino = 4
    
    For i = 2 To ultLinha
    
        Dim livre As Double, cq As Double, bloq As Double
        
        livre = wsOrigem.Cells(i, 4).Value
        cq = wsOrigem.Cells(i, 5).Value
        bloq = wsOrigem.Cells(i, 6).Value
        
        If livre <> 0 Or cq <> 0 Or bloq <> 0 Then
            
            wsNovo.Cells(linhaDestino, 1).Value = wsOrigem.Cells(i, 1).Value
            wsNovo.Cells(linhaDestino, 2).Value = wsOrigem.Cells(i, 2).Value
            wsNovo.Cells(linhaDestino, 3).Value = Format(wsOrigem.Cells(i, 3).Value, "0000000000")
            
            wsNovo.Cells(linhaDestino, 4).Value = livre
            wsNovo.Cells(linhaDestino, 5).Value = cq
            wsNovo.Cells(linhaDestino, 6).Value = bloq
            
            linhaDestino = linhaDestino + 1
            
        End If
        
    Next i
    
    ' =========================
    ' FORMATAÇÃO
    ' =========================
    
    wsNovo.Range("A4:A" & linhaDestino - 1).HorizontalAlignment = xlCenter
    wsNovo.Range("B4:B" & linhaDestino - 1).HorizontalAlignment = xlLeft
    wsNovo.Range("C4:F" & linhaDestino - 1).HorizontalAlignment = xlCenter
    
    wsNovo.Range("D4:F" & linhaDestino - 1).NumberFormat = "#,##0.000"
    
    wsNovo.Columns("A").ColumnWidth = 15
    wsNovo.Columns("B").ColumnWidth = 40
    wsNovo.Columns("C").ColumnWidth = 15
    wsNovo.Columns("D").ColumnWidth = 18
    wsNovo.Columns("E").ColumnWidth = 22
    wsNovo.Columns("F").ColumnWidth = 15
    
    With wsNovo.Range("A3:F" & linhaDestino - 1).Borders
        .LineStyle = xlContinuous
    End With
    
    ' =========================
    ' SALVAR ARQUIVO
    ' =========================
    
    Dim caminho As String
    caminho = ThisWorkbook.Path & "\Posicao_Estoque_" & Format(Now, "yyyymmdd_hhmm") & ".xlsx"
    
    wbNovo.SaveAs caminho, FileFormat:=xlOpenXMLWorkbook
    
    ' =========================
    ' ABRIR AUTOMATICAMENTE ??
    ' =========================
    
    wbNovo.Activate
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "POSIÇÃO DE ESTOQUE GERADA!", vbInformation

End Sub


Sub AbrirAba(nomeAba As String)

    Application.ScreenUpdating = False

    ' ?? GARANTE SENHA ATUAL
    If SENHA = "" Then SENHA = GetSenha

    ' ?? tenta liberar estrutura com segurança
    If ThisWorkbook.ProtectStructure Then
        
        On Error Resume Next
        ThisWorkbook.Unprotect Password:=SENHA
        On Error GoTo 0
        
    End If

    ' ?? se ainda estiver protegido, aborta (evita erro 1004)
    If ThisWorkbook.ProtectStructure Then
        MsgBox "Erro ao liberar estrutura do sistema.", vbCritical
        GoTo Fim
    End If

    With Sheets(nomeAba)
        .Visible = xlSheetVisible
        .Activate
    End With

    ' ?? protege novamente
    ThisWorkbook.Protect Password:=SENHA, Structure:=True

Fim:
    Application.ScreenUpdating = True

End Sub

Sub LimparFormatacaoRange(ws As Worksheet, endereco As String)

    On Error GoTo Finalizar

    Call OtimizarON

    ' Desprotege
    ws.Unprotect SENHA

    Dim intervalo As Range
    Set intervalo = ws.Range(endereco)

    With intervalo
        .Interior.ColorIndex = xlNone
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Color = RGB(0, 0, 0)
    End With

    ' Protege novamente
    ws.Protect SENHA, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True

Finalizar:
    If Err.Number <> 0 Then
    
        MsgBox _
            "Erro " & Err.Number & vbCrLf & vbCrLf & _
            Err.Description & vbCrLf & vbCrLf & _
            "Sub: LimparFormatacaoRange", _
            vbCritical
    
    End If
    Call OtimizarOFF

End Sub

Sub BtnLimparEntradas()
    Call LimparFormatacaoRange(ThisWorkbook.Sheets("Entradas"), "A2:D40")
End Sub

Sub BtnLimparBaixas()
    Call LimparFormatacaoRange(ThisWorkbook.Sheets("Baixas"), "A2:D40")
End Sub

Sub BtnLimparBloqueados()
    Call LimparFormatacaoRange(ThisWorkbook.Sheets("Bloqueados"), "A2:D10")
End Sub

Sub BtnLimparTransferencia()
    Call LimparFormatacaoRange(ThisWorkbook.Sheets("Transferência"), "B6:E60")
End Sub

Sub BtnLimparCQ()
    Call LimparFormatacaoRange(ThisWorkbook.Sheets("Controle de Qualidade"), "A2:D10")
End Sub

Sub BtnLimparEstorno()
    Call LimparFormatacaoRange(ThisWorkbook.Sheets("Estornar Baixa"), "A2:D40")
End Sub


Sub AlterarSenhaSistema()

    Dim ws As Worksheet
    Dim frm As New frmAlterarSenha
    Dim senhaAntigaBackup As String
    
    senhaAntigaBackup = "123"

    SENHA = GetSenha()

    frm.Show
    If frm.Cancelado Then Exit Sub

    ' valida senha atual
    If frm.SenhaAtual <> SENHA Then
        MsgBox "Senha atual incorreta!", vbCritical
        Exit Sub
    End If

    If frm.NovaSenha = "" Then Exit Sub

    If frm.NovaSenha <> frm.ConfirmarSenha Then
        MsgBox "As senhas não conferem!", vbExclamation
        Exit Sub
    End If

    Call OtimizarON

    ' ?? DESPROTEGE WORKBOOK (IMPORTANTE)
    On Error Resume Next
    ThisWorkbook.Unprotect SENHA
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect senhaAntigaBackup
    End If
    On Error GoTo 0

    ' ?? DESPROTEGE PLANILHAS
    For Each ws In ThisWorkbook.Worksheets

        On Error Resume Next
        
        ws.Unprotect SENHA
        
        If ws.ProtectContents Then
            ws.Unprotect senhaAntigaBackup
        End If
        
        On Error GoTo 0

    Next ws

    ' ?? salva nova senha
    ThisWorkbook.Sheets("Configurações").Range("B4").Value = frm.NovaSenha

    SENHA = frm.NovaSenha

    ' ?? PROTEGE WORKBOOK COM NOVA SENHA
    ThisWorkbook.Protect Password:=SENHA, Structure:=True

    ' ?? PROTEGE PLANILHAS
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect SENHA, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    Next ws

    Call OtimizarOFF

    MsgBox "Senha alterada com sucesso!", vbInformation

End Sub
Sub AbrirConfiguracoes()

    Dim senhaAcesso As String
    Dim frm As New frmSenha

    SENHA = GetSenha()

    frm.Show
    If frm.Cancelado Then Exit Sub

    senhaAcesso = frm.senhaDigitada

    If senhaAcesso <> SENHA Then
        MsgBox "Acesso negado!", vbCritical
        Exit Sub
    End If

    Call OtimizarON

    ' ?? FORÇA DESPROTEÇÃO REAL
    On Error Resume Next
    
    ThisWorkbook.Unprotect SENHA
    
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect "123" ' fallback
    End If
    
    On Error GoTo 0

    ' ?? VALIDA SE REALMENTE DESPROTEGEU
    If ThisWorkbook.ProtectStructure Then
        MsgBox "Erro ao desbloquear estrutura do arquivo!", vbCritical
        GoTo Finalizar
    End If

    ' ? AGORA SIM pode mexer
    With ThisWorkbook.Sheets("Configurações")
        .Visible = xlSheetVisible
        .Activate
    End With

Finalizar:
    Call OtimizarOFF

End Sub
Sub FecharConfiguracoes()

    Call OtimizarON

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Configurações")

    ws.Visible = xlVeryHidden

    Sheets("Dashboard").Activate

    ' trava estrutura de novo
    Call ProtegerSistema

    Call OtimizarOFF

End Sub


Function GetSenha() As String

    On Error Resume Next
    
    Dim s As String
    s = ThisWorkbook.Sheets("Configurações").Range("B4").Value

    If s = "" Then s = "123"

    SENHA = s
    GetSenha = s

End Function


Sub DesprotegerSistema()

    Dim ws As Worksheet

    On Error Resume Next

    ' Workbook
    ThisWorkbook.Unprotect SENHA
    ThisWorkbook.Unprotect "123"

    ' Sheets
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect SENHA
        ws.Unprotect "123"
    Next ws

    On Error GoTo 0

End Sub

Sub ProtegerSistema()

    Dim ws As Worksheet

    SENHA = GetSenha()

    For Each ws In ThisWorkbook.Worksheets

        ws.Protect _
            Password:=SENHA, _
            UserInterfaceOnly:=True, _
            AllowFiltering:=True, _
            AllowSorting:=True

    Next ws

    ThisWorkbook.Protect Password:=SENHA, Structure:=True

End Sub

Sub CancelarAutoSave()

    On Error Resume Next
    
    Application.OnTime ProximoAutoSave, "ExecutarAutoSave", , False

End Sub

Sub CarregarMateriaisCache()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Materiais")

    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set dictMateriais = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim cod As String
    Dim desc As String

    For i = 2 To ult

        cod = Trim(CStr(ws.Cells(i, 1).Value))
        desc = Trim(CStr(ws.Cells(i, 2).Value))

        If cod <> "" Then
            dictMateriais(cod) = desc
        End If

    Next i

End Sub

Function BuscarDescricaoRapida(cod As String) As String

    On Error GoTo Finalizar

    If dictMateriais Is Nothing Then
        Call CarregarMateriaisCache
    End If

    cod = Trim(CStr(cod))

    If dictMateriais.Exists(cod) Then
        
        BuscarDescricaoRapida = dictMateriais(cod)
        
    Else
        
        BuscarDescricaoRapida = "[NÃO CADASTRADO]"
        
    End If

    Exit Function

Finalizar:

    BuscarDescricaoRapida = "[ERRO]"

End Function

'###########SUB PARA CRIAR NOVAS ABAS###########
Sub CriarAbaADM()

    Dim ws As Worksheet
    
    Application.ScreenUpdating = False

    ' ?? garante senha atual
    If SENHA = "" Then SENHA = GetSenha

    ' ?? libera estrutura
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect Password:=SENHA
    End If

    ' ?? verifica se já existe
    On Error Resume Next
    Set ws = Sheets("ADM")
    On Error GoTo 0

    If ws Is Nothing Then
        
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "ADM"
        
        ' ?? opcional: já deixa padrão
        ws.Cells(1, 1).Value = "ÁREA ADMINISTRATIVA"
        ws.Cells.Font.Name = "Arial"
        ws.Cells.Font.Size = 11
        
    Else
        MsgBox "A aba ADM já existe.", vbInformation
    End If

    ' ?? protege de novo
    ThisWorkbook.Protect Password:=SENHA, Structure:=True

    Application.ScreenUpdating = True

End Sub

Sub AbrirADM()

    Dim frm As New frmSenha
    Dim senhaAcesso As String

    ' ?? garante senha atual do sistema
    If SENHA = "" Then SENHA = GetSenha

    ' ?? abre form (senha mascarada)
    frm.Show

    ' ? cancelou
    If frm.Cancelado Then Exit Sub

    senhaAcesso = frm.senhaDigitada

    ' ?? valida senha
    If senhaAcesso <> SENHA Then
        MsgBox "Acesso negado!", vbCritical
        Exit Sub
    End If

    ' ? acesso liberado
    Call AbrirAba("ADM")

End Sub

Sub OrdenarHistorico()

    On Error GoTo Fim

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Historico")

    Dim ult As Long
    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If ult < 3 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    With ws.Sort

        .SortFields.Clear

        ' DATA/HORA MAIS RECENTE
        .SortFields.Add _
            Key:=ws.Range("G2:G" & ult), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal

        ' CÓDIGO MATERIAL
        .SortFields.Add _
            Key:=ws.Range("A2:A" & ult), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SetRange ws.Range("A1:J" & ult)

        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom

        .Apply

    End With

Fim:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Sub ExportarEstoqueConsolidado_Excel()

    On Error GoTo TratarErro

    Dim wsOrigem As Worksheet
    Dim wbNovo As Workbook
    Dim wsNovo As Worksheet

    Dim dict As Object
    Dim arr As Variant
    Dim chave As Variant

    Dim ultLinha As Long
    Dim linhaDestino As Long
    Dim i As Long

    Dim cod As String
    Dim desc As String

    Dim livre As Double
    Dim cq As Double
    Dim bloq As Double

    Dim caminho As String

    Set wsOrigem = ThisWorkbook.Sheets("Estoque Atual")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' =========================
    ' DICIONÁRIO
    ' =========================

    Set dict = CreateObject("Scripting.Dictionary")

    ultLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultLinha

        cod = Trim(wsOrigem.Cells(i, 1).Value)

        If cod <> "" Then

            desc = wsOrigem.Cells(i, 2).Value

            livre = Val(wsOrigem.Cells(i, 4).Value)
            cq = Val(wsOrigem.Cells(i, 5).Value)
            bloq = Val(wsOrigem.Cells(i, 6).Value)

            If Not dict.Exists(cod) Then

                dict.Add cod, Array(desc, livre, cq, bloq)

            Else

                arr = dict(cod)

                arr(1) = arr(1) + livre
                arr(2) = arr(2) + cq
                arr(3) = arr(3) + bloq

                dict(cod) = arr

            End If

        End If

    Next i

    ' =========================
    ' NOVO ARQUIVO
    ' =========================

    Set wbNovo = Workbooks.Add
    Set wsNovo = wbNovo.Sheets(1)

    ' nome MENOR que 31 caracteres
    wsNovo.Name = "Estoque Consolidado"

    ' =========================
    ' CABEÇALHO
    ' =========================

    wsNovo.Range("A1:F1").Merge

    wsNovo.Range("A1").Value = _
        "RELATÓRIO CONSOLIDADO DE ESTOQUE - " & _
        Format(Now, "dd/mm/yyyy hh:mm")

    With wsNovo.Range("A1")

        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter

    End With

    ' =========================
    ' TÍTULOS
    ' =========================

    wsNovo.Range("A3").Value = "Material"
    wsNovo.Range("B3").Value = "Descrição"
    wsNovo.Range("C3").Value = "Livre"
    wsNovo.Range("D3").Value = "CQ"
    wsNovo.Range("E3").Value = "Bloqueado"
    wsNovo.Range("F3").Value = "Total"

    With wsNovo.Range("A3:F3")

        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous

    End With

    ' =========================
    ' DADOS
    ' =========================

    linhaDestino = 4

    For Each chave In dict.Keys

        arr = dict(chave)

        wsNovo.Cells(linhaDestino, 1).Value = chave
        wsNovo.Cells(linhaDestino, 2).Value = arr(0)

        wsNovo.Cells(linhaDestino, 3).Value = arr(1)
        wsNovo.Cells(linhaDestino, 4).Value = arr(2)
        wsNovo.Cells(linhaDestino, 5).Value = arr(3)

        wsNovo.Cells(linhaDestino, 6).Value = _
            arr(1) + arr(2) + arr(3)

        linhaDestino = linhaDestino + 1

    Next chave

    ' =========================
    ' FORMATAÇÃO
    ' =========================

    wsNovo.Range("C4:F" & linhaDestino - 1).NumberFormat = "#,##0.000"

    wsNovo.Columns("A").ColumnWidth = 15
    wsNovo.Columns("B").ColumnWidth = 45
    wsNovo.Columns("C").ColumnWidth = 15
    wsNovo.Columns("D").ColumnWidth = 15
    wsNovo.Columns("E").ColumnWidth = 15
    wsNovo.Columns("F").ColumnWidth = 15

    With wsNovo.Range("A3:F" & linhaDestino - 1).Borders
        .LineStyle = xlContinuous
    End With

    ' =========================
    ' SALVAR
    ' =========================

    caminho = ThisWorkbook.Path & _
        "\Estoque_Consolidado_" & _
        Format(Now, "yyyymmdd_hhmm") & ".xlsx"

    wbNovo.SaveAs caminho, FileFormat:=xlOpenXMLWorkbook

    wbNovo.Activate

    MsgBox "RELATÓRIO DE ESTOQUE CONSOLIDADO GERADO!", vbInformation

Finalizar:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Exit Sub

TratarErro:

    MsgBox _
        "Erro " & Err.Number & vbCrLf & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Sub: ExportarEstoqueConsolidado_Excel", _
        vbCritical

    Resume Finalizar

End Sub

