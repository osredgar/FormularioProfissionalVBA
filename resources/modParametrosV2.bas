Attribute VB_Name = "modParametros"
Option Explicit
Option Private Module

'Private Const SENHA_PLAN As String = "gestaocriativa123*"
Private Const SENHA_PLAN As String = "teste123"

Public Sub DesbloquearTodasPlanilhas()
    AtivarConfiguracoesExcel False
    DesbloquearPastaTrabalho
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        DesbloquearPlanilha ws, False
        ws.Activate
        ws.Visible = xlSheetVisible
        ActiveWindow.DisplayGridlines = True
        ActiveWindow.DisplayHeadings = True
    Next ws
    AtivarConfiguracoesExcel True
    Set ws = Nothing
    shHome.Activate
End Sub

Public Sub BloquearTodasPlanilhas()
    AtivarConfiguracoesExcel False
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        Select Case ws.CodeName
            Case "shBancoDados", "shAuxiliar"
                ws.Visible = xlSheetVeryHidden
            Case Else
                BloquearPlanilha ws, False
        End Select
    Next ws
    
    BloquearPastaTrabalho
    
    AtivarConfiguracoesExcel True
    
    shHome.Activate
End Sub

Public Sub DesbloquearPastaTrabalho()
    ThisWorkbook.Unprotect SENHA_PLAN
End Sub

Public Sub BloquearPastaTrabalho()
    ThisWorkbook.Protect SENHA_PLAN, True, True
End Sub

Public Sub DesbloquearPlanilha(ByRef ws As Worksheet, Optional ByVal desbloquearPasta As Boolean = True)
    If desbloquearPasta Then DesbloquearPastaTrabalho
    ws.Unprotect SENHA_PLAN
End Sub

Public Sub BloquearPlanilha(ByRef ws As Worksheet, _
                            Optional bloquearPasta As Boolean = True, _
                            Optional somenteInterface As Boolean = True, _
                            Optional podeFiltrar As Boolean = True, _
                            Optional podeExcluir As Boolean = False, _
                            Optional podeIncluir As Boolean = False, _
                            Optional podeFormatar As Boolean = False)

    ws.Protect Password:=SENHA_PLAN, _
               AllowDeletingRows:=podeExcluir, _
               AllowDeletingColumns:=podeExcluir, _
               AllowInsertingRows:=podeIncluir, _
               AllowInsertingColumns:=podeIncluir, _
               AllowFormattingCells:=podeFormatar, _
               AllowFiltering:=podeFiltrar, _
               AllowSorting:=podeFiltrar, _
               Contents:=somenteInterface, _
               Scenarios:=somenteInterface, _
               UserInterfaceOnly:=somenteInterface, _
               AllowUsingPivotTables:=somenteInterface
    
    If bloquearPasta Then BloquearPastaTrabalho
End Sub

Public Sub AtivarConfiguracoesExcel(ByVal ativar As Boolean)
    With Application
        .ScreenUpdating = ativar
        .EnableEvents = ativar
        .DisplayAlerts = ativar
        Select Case ativar
            Case True
                .Calculation = xlCalculationAutomatic
                .StatusBar = False
            Case False
                .Calculation = xlCalculationManual
                .StatusBar = "Processando dados aguarde..."
        End Select
    End With
End Sub


Public Sub MostrarAbasPlanilhas(ByVal mostrar As Boolean)
    With ActiveWindow
        Let .DisplayGridlines = False
        Let .DisplayHeadings = False
        Let .DisplayHorizontalScrollBar = False
        Let .DisplayVerticalScrollBar = False
        Let .DisplayWorkbookTabs = mostrar
        Let .EnableResize = False
        Let .DisplayFormulas = False
    End With
    Let Application.DisplayStatusBar = mostrar
End Sub

'**********************************************************************
'DEFINIR TECLA DE ATALHO PARA EXECUTAR ALGUM PROCEDIMENTO
'PARÂMETROS: NOME DA TECLA, NOME DO PROCEDIMENTO
'**********************************************************************
Public Sub DefinirAtalho(ByVal tecla As String, ByVal procedimento As String)
    Call Application.OnKey(tecla, procedimento)
End Sub

'**********************************************************************
'REMOVER PROCEDIMENTO DA TECLA DE ATALHO
'PARÂMETROS: NOME DA TECLA
'**********************************************************************
Public Sub RemoverAtalho(ByVal tecla As String)
    Call Application.OnKey(tecla, vbNullString)
End Sub

Public Sub SalvarFecharExcel()
    ' Desativa eventos para evitar que outros eventos sejam acionados
    Application.EnableEvents = False
    
    ' Salva a pasta de trabalho atual sem solicitar confirmação
    ThisWorkbook.SaveAs ThisWorkbook.FullName
    
    ' Ativa eventos novamente para restaurar o comportamento normal do Excel
    Application.EnableEvents = True
    
    ' Fecha a pasta de trabalho atual sem salvar alterações
    ThisWorkbook.Close SaveChanges:=False
    
    ' Fecha o Excel completamente
    Application.Quit
End Sub


Public Sub ExibirTelaCheia()
    With Application
        .DisplayFullScreen = Not .DisplayFullScreen
    End With
    ActiveSheet.Cells(1, 1).Select
End Sub


Public Sub ConfigurarTela(ByVal exibir As Boolean)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        With ws
            .Activate
            .DisplayPageBreaks = False
            .Cells(1, 1).Select
        End With
        With ActiveWindow
            .DisplayGridlines = exibir
            .DisplayHeadings = exibir
            .DisplayFormulas = exibir
            .DisplayVerticalScrollBar = exibir
            .DisplayHorizontalScrollBar = exibir
            .DisplayWorkbookTabs = exibir
        End With
    Next ws
End Sub
