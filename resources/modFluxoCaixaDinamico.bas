Attribute VB_Name = "modFluxoCaixaDinamico"
Option Explicit
Option Private Module


Private Function tbLancamentosDinamico() As ListObject
    Set tbLancamentosDinamico = shAuxDinamicas.ListObjects("tbLancamentosDinamico")
End Function

Private Function tbEstruturasDinamico() As ListObject
    Set tbEstruturasDinamico = shAuxDinamicas.ListObjects("tbEstruturasDinamico")
End Function

Private Function tbDinFluxoConsolidado() As PivotTable
    Set tbDinFluxoConsolidado = shDinFluxoConsolidado.PivotTables("tbDinFluxoConsolidado")
End Function

Private Function tbDinFluxoResumido() As PivotTable
    Set tbDinFluxoResumido = shDinFluxoResumido.PivotTables("tbDinFluxoResumido")
End Function

Private Function tbDinFluxoDetalhado() As PivotTable
    Set tbDinFluxoDetalhado = shDinFluxoDetalhado.PivotTables("tbDinFluxoDetalhado")
End Function

Private Function tbDinMedidas() As PivotTable
    Set tbDinMedidas = shAuxDashboard.PivotTables("tbDinMedidas")
End Function

Private Function tbDinGraficoEvolucao() As ChartObject
    Set tbDinGraficoEvolucao = shDinDashboard.ChartObjects("tbDinGraficoEvolucao")
End Function

Private Function tbDinGraficoContas() As ChartObject
    Set tbDinGraficoContas = shDinDashboard.ChartObjects("tbDinGraficoContas")
End Function

Public Sub FluxoDinamico_AlterarBotaoSelecionado(ByVal botao As Object, ByVal selecionado As Boolean)
    With botao

        .BorderColor = RGB(118, 116, 118)

        Select Case selecionado
            Case True
                .BackColor = RGB(118, 116, 118)
                .ForeColor = RGB(255, 255, 255)

            Case False
                .BackColor = RGB(255, 255, 255)
                .ForeColor = RGB(118, 116, 118)
        End Select

    End With
End Sub

Public Sub LancamentosDinamico_Carregar()
    AtivarConfiguracoesExcel False
    
    AtualizarBarraProgresso 1 / 12, "Carregando Lançamentos"
    
    Misc_LimparTabelaDataBodyRange tbLancamentosDinamico
    With tbLancamentosDinamico
        Dim sql As String: sql = "SELECT * FROM vwLancamentosDinamico;"
        PreencherPlanilha .Parent, sql, .Range(1, 1).Address, True, False, False, False
    End With
    
    AtualizarBarraProgresso 2 / 12, "Carregando Estruturas"
    
    Misc_LimparTabelaDataBodyRange tbEstruturasDinamico
    With tbEstruturasDinamico
        sql = "SELECT * FROM vwEstruturasDinamico;"
        PreencherPlanilha .Parent, sql, .Range(1, 1).Address, True, False, False, False
    End With
    
    AtualizarBarraProgresso 3 / 12, "Atualizando Conexões - Lançamentos"
    ThisWorkbook.Connections("Consulta - lancamentos").Refresh
    
    AtualizarBarraProgresso 4 / 12, "Atualizando Conexões - Estruturas"
    ThisWorkbook.Connections("Consulta - estruturas").Refresh
    
    AtualizarBarraProgresso 5 / 12, "Atualizando Conexões - Datas"
    ThisWorkbook.Connections("Consulta - datas").Refresh
    
    AtualizarBarraProgresso 6 / 12, "Atualizando Modelo de Dados"
    ThisWorkbook.Connections("ThisWorkBookDataModel").Refresh
    Application.Wait (Now + TimeValue("0:00:05"))
    
    AtualizarBarraProgresso 7 / 12, "Atualizando Fluxo Consolidado"
    With tbDinFluxoConsolidado
        DesbloquearPlanilha .Parent
        .RefreshTable
        BloquearPlanilha .Parent
    End With
    
    AtualizarBarraProgresso 8 / 12, "Atualizando Fluxo Detalhado"
    With tbDinFluxoDetalhado
        DesbloquearPlanilha .Parent
        .PivotCache.Refresh
        .RefreshTable
        BloquearPlanilha .Parent
    End With

    AtualizarBarraProgresso 9 / 12, "Atualizando Fluxo Resumido"
    With tbDinFluxoResumido
        DesbloquearPlanilha .Parent
        .PivotCache.Refresh
        .RefreshTable
        BloquearPlanilha .Parent
    End With
    
    AtualizarBarraProgresso 10 / 12, "Atualizando Dashboard"
    With tbDinMedidas
        .PivotCache.Refresh
        .RefreshTable
        Dim dataAtualizacao As String: dataAtualizacao = .PivotCache.RefreshDate
    End With
    
    AtualizarBarraProgresso 11 / 12, "Limpando tabelas temporárias"
    Misc_LimparTabelaDataBodyRange tbEstruturasDinamico
    Misc_LimparTabelaDataBodyRange tbLancamentosDinamico

    AtualizarBarraProgresso 12 / 12, "Finalizando"
    FecharBarraProgresso
        
    AtivarConfiguracoesExcel True
    
    Detalhamento_Carregar
    
    MsgBox "Visual atualizado:" & vbCrLf & dataAtualizacao, vbInformation, APP_NOME
End Sub

Public Sub FluxoConsolidado_LimparFiltros()
    Misc_TabelaDinamicaLimparTodosFiltros tbDinFluxoConsolidado
End Sub

Public Sub FluxoConsolidado_AlterarVisualSelecionado(ByVal botao As Object)
    With tbDinFluxoConsolidado
        On Error Resume Next
        
        Dim cfMensal As CubeField: Set cfMensal = .CubeFields("[datas].[MesNome]")
        Dim cfSemanal As CubeField: Set cfSemanal = .CubeFields("[datas].[Semana]")
        Dim cfDiario As CubeField: Set cfDiario = .CubeFields("[datas].[Dia]")
        Dim cfOrcado As CubeField: Set cfOrcado = .CubeFields("[Measures].[fxFluxoCaixaPrevisto]")
        Dim cfRealizado As CubeField: Set cfRealizado = .CubeFields("[Measures].[fxFluxoCaixaRealizado]")
        
        cfMensal.Orientation = xlColumnField
        cfSemanal.Orientation = xlHidden
        cfDiario.Orientation = xlHidden
        cfRealizado.Orientation = xlDataField

        Dim pfMeses As PivotField: Set pfMeses = .PivotFields("[datas].[MesNome].[MesNome]")
        
        Dim selecaoMesAtual As Variant: selecaoMesAtual = Array(pfMeses.PivotItems(pfMeses.PivotItems.Count))
        
        On Error GoTo 0
        
        Select Case botao.Caption
            Case "Mensal"
                cfOrcado.Orientation = xlDataField
                cfOrcado.Position = 1
                pfMeses.ClearAllFilters
                
            Case "Semanal"
                pfMeses.VisibleItemsList = selecaoMesAtual
                cfOrcado.Orientation = xlHidden
                cfMensal.Orientation = xlHidden
                cfSemanal.Orientation = xlColumnField
                
            Case "Diário"
                pfMeses.VisibleItemsList = selecaoMesAtual
                cfOrcado.Orientation = xlHidden
                cfMensal.Orientation = xlHidden
                cfDiario.Orientation = xlColumnField
                
            Case "Orçado"
                pfMeses.ClearAllFilters
                Select Case True
                    Case cfOrcado.Orientation = xlHidden
                        FluxoDinamico_AlterarBotaoSelecionado botao, True
                        cfOrcado.Orientation = xlDataField
                        cfOrcado.Position = 1
                        
                    Case cfOrcado.Orientation = xlDataField
                        FluxoDinamico_AlterarBotaoSelecionado botao, False
                        cfOrcado.Orientation = xlHidden
                    
                End Select
        End Select

        'Aplicar a formatação dos números
        On Error Resume Next
        .DataBodyRange.NumberFormatLocal = FORMATO_DECIMAL
        On Error GoTo 0
    End With
    
    CorrigirVisualConsolidado
End Sub

Public Sub FluxoResumido_LimparFiltros()
    Misc_TabelaDinamicaLimparTodosFiltros tbDinFluxoResumido
End Sub

Public Sub DashboardDinamico_LimparFiltros()
    Misc_TabelaDinamicaLimparTodosFiltros tbDinMedidas
    DesbloquearPlanilha shAuxDashboard
End Sub

Public Sub FluxoDetalhado_LimparFiltros()
    Misc_TabelaDinamicaLimparTodosFiltros tbDinFluxoDetalhado
End Sub

Public Sub FluxoDetalhado_AlterarVisualSelecionado(ByVal botao As Object)
     With tbDinFluxoDetalhado
        On Error Resume Next
        
        Dim cfMensal As CubeField: Set cfMensal = .CubeFields("[datas].[MesNome]")
        Dim cfSemanal As CubeField: Set cfSemanal = .CubeFields("[datas].[Semana]")
        Dim cfDiario As CubeField: Set cfDiario = .CubeFields("[datas].[Dia]")
        Dim cfOrcado As CubeField: Set cfOrcado = .CubeFields("[Measures].[fxFluxoCaixaPrevisto]")
        Dim cfRealizado As CubeField: Set cfRealizado = .CubeFields("[Measures].[fxFluxoCaixaRealizado]")

        cfMensal.Orientation = xlColumnField
        cfSemanal.Orientation = xlHidden
        cfDiario.Orientation = xlHidden
        cfRealizado.Orientation = xlDataField
        
        Dim pfMeses As PivotField: Set pfMeses = .PivotFields("[datas].[MesNome].[MesNome]")
        
        Dim selecaoMesAtual As Variant: selecaoMesAtual = Array(pfMeses.PivotItems(pfMeses.PivotItems.Count))
        
        On Error GoTo 0
        
        Select Case botao.Caption
            Case "Mensal"
                cfOrcado.Orientation = xlDataField
                cfOrcado.Position = 1
                pfMeses.ClearAllFilters
                
            Case "Semanal"
                pfMeses.VisibleItemsList = selecaoMesAtual
                cfOrcado.Orientation = xlHidden
                cfMensal.Orientation = xlHidden
                cfSemanal.Orientation = xlColumnField
                
            Case "Diário"
                pfMeses.VisibleItemsList = selecaoMesAtual
                cfOrcado.Orientation = xlHidden
                cfMensal.Orientation = xlHidden
                cfDiario.Orientation = xlColumnField
                
            Case "Orçado"
                pfMeses.ClearAllFilters
                Select Case True
                    Case cfOrcado.Orientation = xlHidden
                        FluxoDinamico_AlterarBotaoSelecionado botao, True
                        cfOrcado.Orientation = xlDataField
                        cfOrcado.Position = 1
                        
                    Case cfOrcado.Orientation = xlDataField
                        FluxoDinamico_AlterarBotaoSelecionado botao, False
                        cfOrcado.Orientation = xlHidden
                    
                End Select
        End Select

        'Aplicar a formatação dos números
        .DataBodyRange.NumberFormatLocal = FORMATO_DECIMAL
    End With
    
    CorrigirVisualDetalhado
End Sub

'Private Sub DepurarCubeFields()
'    Dim cf As CubeField
'    For Each cf In shDinFluxoConsolidado.PivotTables(1).CubeFields
'        With cf
'            Debug.Print "Name: " & .Name, "Orientation: " & .Orientation
'            If .Orientation <> xlHidden Then
'                Debug.Print "Position: " & .Position
'            End If
'        End With
'    Next cf
'
'    Debug.Print String(80, "=")
'
'    Dim pf As PivotField
'    For Each pf In shDinFluxoDetalhado.PivotTables(1).DataFields
'        Debug.Print "SourceName:" & pf.SourceName & vbCrLf & "Name:" & pf.Name & vbCrLf & "NumberFormat:" & pf.NumberFormat
'    Next pf
'End Sub

Public Sub FluxoCaixa_GerarArquivoPDF(ByVal planilhas As Variant)
    Dim caminhoPDF As String: caminhoPDF = _
        Application.GetSaveAsFilename(FileFilter:="PDF Files (*.pdf), *.pdf", Title:="Salvar Como PDF")
    
    If caminhoPDF = "" Or caminhoPDF = "False" Then
        MsgBox "Nenhuma pasta selecionada, operação cancelada pelo usuário", vbCritical, APP_NOME
        Exit Sub
    End If
    
    If Not Right(caminhoPDF, 4) = ".pdf" Then
        caminhoPDF = caminhoPDF & ".pdf"
    End If
    
    ' Exibe a pré-visualização
    ThisWorkbook.Sheets(planilhas).PrintPreview
    
    ' Exporta apenas as planilhas selecionadas como PDF
    ThisWorkbook.Sheets(planilhas).Select
    
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=caminhoPDF, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "PDF salvo com sucesso em:" & vbCrLf & caminhoPDF, vbInformation, APP_NOME
End Sub

Private Sub CorrigirVisualDetalhado()
    On Error Resume Next
    With tbDinFluxoDetalhado.PivotFields("[estruturas].[GrupoContas].[GrupoContas]")
        .DrilledDown = False
        .PivotItems("[estruturas].[GrupoContas].&[ENTRADAS]").DrilledDown = True
        .PivotItems("[estruturas].[GrupoContas].&[SAÍDAS]").DrilledDown = True
        .PivotItems("[estruturas].[GrupoContas].&[DISTRIBUIÇÃO DE DIVIDENDOS]").DrilledDown = True
    End With
    On Error GoTo 0
End Sub


Private Sub CorrigirVisualConsolidado()
    On Error Resume Next
    With tbDinFluxoConsolidado.PivotFields("[estruturas].[Nivel1].[Nivel1]")
        .DrilledDown = False
        .PivotItems("[estruturas].[Nivel1].&[Fluxo De Caixa Operacional]").DrilledDown = True
        .PivotItems("[estruturas].[Nivel1].&[Fluxo De Caixa Não Operacional]").DrilledDown = True
    End With
    On Error GoTo 0
End Sub

Private Function tbDetalhamento() As ListObject
    Set tbDetalhamento = shDetalhamento.ListObjects("tbDetalhamento")
End Function

Public Sub Detalhamento_Carregar()
    AtivarConfiguracoesExcel False
    
    Dim tabela As ListObject: Set tabela = tbDetalhamento
    
    With tabela
        DesbloquearPlanilha .Parent
        
        AtualizarBarraProgresso 1 / 4, "Limpando registros antigos"
        Misc_LimparTabelaDataBodyRange tabela
        
        AtualizarBarraProgresso 2 / 4, "Carregando novos registros"
        Dim sql As String: sql = "SELECT * FROM vwDetalhamentoCarga;"
        PreencherPlanilha .Parent, sql, .Range(1, 1).Address, True, False, False, False
        
        AtualizarBarraProgresso 3 / 4, "Aplicando formatações"
        .ListColumns("Valor").DataBodyRange.NumberFormatLocal = FORMATO_DECIMAL
        .Parent.Range("B4:B5").NumberFormatLocal = FORMATO_DECIMAL
        
        BloquearPlanilha .Parent
    End With
    
    AtivarConfiguracoesExcel True

    AtualizarBarraProgresso 4 / 4, "Finalizando"
    FecharBarraProgresso
End Sub

Public Sub Detalhamento_LimparFiltros()
    Misc_LimparFiltrosTabela tbDetalhamento
End Sub
