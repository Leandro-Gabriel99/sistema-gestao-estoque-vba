VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    On Error GoTo TratarErro

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim etapaAtual As String
    Dim ws As Worksheet

    ' =========================
    ' SENHA
    ' =========================
    etapaAtual = "GetSenha"

    SENHA = GetSenha()

    ' =========================
    ' OCULTAR GUIAS
    ' =========================
    etapaAtual = "Ocultar guias"

    Application.ActiveWindow.DisplayWorkbookTabs = False

    ' =========================
    ' DESPROTEGE ESTRUTURA
    ' =========================
    etapaAtual = "Desproteger estrutura"

    ThisWorkbook.Unprotect Password:=SENHA

    ' =========================
    ' PROTEÇÃO DAS PLANILHAS
    ' =========================
    etapaAtual = "Proteger planilhas"

    For Each ws In ThisWorkbook.Worksheets

        Select Case ws.Name

            Case "Estoque Atual", "Historico", "Entradas", "Baixas", _
                 "Bloqueados", "Consultar Movimentação", _
                 "Controle de Qualidade", "Estornar Baixa", _
                 "Transferência", "Busca Materiais", "TEMP_PDF"

                ws.Unprotect Password:=SENHA
                ws.Protect Password:=SENHA, UserInterfaceOnly:=True

        End Select

    Next ws

    ' =========================
    ' GARANTE DASHBOARD VISÍVEL
    ' =========================
    etapaAtual = "Mostrar Dashboard"

    Sheets("Dashboard").Visible = xlSheetVisible

    ' =========================
    ' OCULTA AS DEMAIS
    ' =========================
    etapaAtual = "Ocultar abas"

    For Each ws In ThisWorkbook.Worksheets

        If ws.Name <> "Dashboard" Then
            ws.Visible = xlVeryHidden
        End If

    Next ws

    ' =========================
    ' ABRE DASHBOARD
    ' =========================
    etapaAtual = "Abrir Dashboard"

    Sheets("Dashboard").Activate

    ' =========================
    ' RECOLORIR HISTÓRICO
    ' =========================
    etapaAtual = "RecolorirHistorico"

    Call RecolorirHistorico

    ' =========================
    ' PROTEGER SISTEMA
    ' =========================
    etapaAtual = "ProtegerSistema"

    Call ProtegerSistema

    ' =========================
    ' AUTOSAVE
    ' =========================
    etapaAtual = "AutoSaveLoop"

    Call AutoSaveLoop

    ' =========================
    ' CACHE DE MATERIAIS
    ' =========================
    etapaAtual = "CarregarMateriaisCache"

    Call CarregarMateriaisCache

    ' =========================
    ' REPROTEGE ESTRUTURA
    ' =========================
    etapaAtual = "Proteger estrutura"

    ThisWorkbook.Protect Password:=SENHA, Structure:=True

Finalizar:

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Exit Sub

TratarErro:

    MsgBox _
        "Erro " & Err.Number & vbCrLf & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Etapa: " & etapaAtual & vbCrLf & _
        "Evento: Workbook_Open", _
        vbCritical

    Resume Finalizar

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error GoTo TratarErro

    Dim etapaAtual As String

    ' =========================
    ' CANCELA AUTOSAVE
    ' =========================
    etapaAtual = "CancelarAutoSave"

    Call CancelarAutoSave

    ' =========================
    ' SALVAR
    ' =========================
    etapaAtual = "SalvarSeguro"

    Application.DisplayAlerts = False

    Call SalvarSeguro

    Application.DisplayAlerts = True

Finalizar:

    Application.DisplayAlerts = True

    Exit Sub

TratarErro:

    MsgBox _
        "Erro " & Err.Number & vbCrLf & vbCrLf & _
        Err.Description & vbCrLf & vbCrLf & _
        "Etapa: " & etapaAtual & vbCrLf & _
        "Evento: Workbook_BeforeClose", _
        vbCritical

    Resume Finalizar

End Sub

