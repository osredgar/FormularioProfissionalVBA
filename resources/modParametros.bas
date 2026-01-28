Attribute VB_Name = "modParametros"
Option Explicit
Option Private Module

Private Const SENHA_PLAN As String = "eosr123*"


Public Sub DesbloquearTodasPlanilhas()
    AtivarConfiguracoesExcel False
    
    DesbloquearPastaTrabalho
    
    MostrarAbasPlanilhas True
        
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        With ws
            .Unprotect SENHA_PLAN
            .Activate
            .Cells(1, 1).Select
            .Visible = xlSheetVisible
            .ScrollArea = ""
            .EnableSelection = xlNoRestrictions
        End With
        
    Next ws
    
    AtivarConfiguracoesExcel True
End Sub

Public Sub BloquearTodasPlanilhas()
    AtivarConfiguracoesExcel False
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        With ws
            .Activate
            .Cells(1, 1).Select
            .EnableSelection = xlNoRestrictions
            
            Select Case .CodeName
                Case "shLogin", "shBancoDados", "shAuxiliar"
                    .Visible = xlSheetVeryHidden
                
                Case Else
                    BloquearPlanilha ws, False

            End Select
        
        End With
    
    Next ws
    
    MostrarAbasPlanilhas False
    
    BloquearPastaTrabalho
    
    AtivarConfiguracoesExcel True
End Sub

Public Sub BloquearPastaTrabalho()
    ThisWorkbook.Protect SENHA_PLAN, True, True
End Sub

Public Sub DesbloquearPastaTrabalho()
    ThisWorkbook.Unprotect SENHA_PLAN
End Sub

Public Sub DesbloquearPlanilha(ByRef ws As Worksheet, Optional ByVal desbloquearPasta As Boolean = True)
    ws.Unprotect SENHA_PLAN
    If desbloquearPasta Then DesbloquearPastaTrabalho
End Sub

Public Sub BloquearPlanilha(ByRef ws As Worksheet, _
                            Optional ByVal bloquearPasta As Boolean = True, _
                            Optional ByVal somenteInterface As Boolean = True, _
                            Optional ByVal podeFiltrar As Boolean = True, _
                            Optional ByVal podeIncluir As Boolean = False, _
                            Optional ByVal podeExcluir As Boolean = False, _
                            Optional ByVal podeFormatar As Boolean = False)
    
    With ws
        .Protect _
            Password:=SENHA_PLAN, _
            DrawingObjects:=somenteInterface, _
            Contents:=somenteInterface, _
            AllowDeletingRows:=podeExcluir, _
            AllowDeletingColumns:=podeExcluir, _
            AllowInsertingRows:=podeIncluir, _
            AllowInsertingColumns:=podeIncluir, _
            AllowFormattingCells:=podeFormatar, _
            AllowFiltering:=podeFiltrar, _
            AllowSorting:=podeFiltrar, _
            Scenarios:=somenteInterface, _
            UserInterfaceOnly:=somenteInterface, _
            AllowUsingPivotTables:=somenteInterface
        
        .EnableSelection = xlNoRestrictions
    End With
    
    If bloquearPasta Then BloquearPastaTrabalho
End Sub

Public Sub AtivarConfiguracoesExcel(ByVal ativar As Boolean)
    With Application
        .DisplayAlerts = ativar
        .ScreenUpdating = ativar
        .EnableEvents = ativar
        .EnableAnimations = ativar
        .CutCopyMode = False
    
        Select Case ativar
            Case True
                .Calculation = xlCalculationAutomatic
                .StatusBar = False
            Case False
                .Calculation = xlCalculationManual
                .StatusBar = "Executando..."
        End Select
    End With
End Sub

Public Sub MostrarAbasPlanilhas(ByVal ativar As Boolean)
    If ActiveWorkbook Is ThisWorkbook Then
        With Application.ActiveWindow
            .DisplayGridlines = ativar
            .DisplayHeadings = ativar
            .DisplayWorkbookTabs = ativar
            .DisplayHorizontalScrollBar = ativar
            .DisplayVerticalScrollBar = ativar
        End With
    End If
End Sub

Public Sub ExibirTelaInteira()
    With Application
        .DisplayFullScreen = Not .DisplayFullScreen
    End With
End Sub


Public Sub DefinirAtalho(ByVal tecla As String, Optional ByVal procedimento As String = "")
'"^{F11}" - Ctrl+F11
'"+{F11}" - Shift+F11
'"%{F11}" - Alt+F11
    Application.OnKey tecla, procedimento
End Sub

Public Sub ExibirJanelaDoExcel(ByVal exibir As Boolean)
    With ThisWorkbook
        .Windows(.Name).Visible = exibir
    End With
End Sub
