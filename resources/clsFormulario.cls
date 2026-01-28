VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim oFormulario As MSForms.UserForm
Dim oTabela As ListObject
Dim oAutoIncremento As ListObject
Dim mIndex As Long
Dim mCodigo As Long


Public Property Set Formulario(uf As MSForms.UserForm)
    Set oFormulario = uf
End Property

Public Property Set Tabela(lo As ListObject)
    Set oTabela = lo
End Property

Public Property Set TabelaIndices(lo As ListObject)
    Set oAutoIncremento = lo
End Property

Private Sub Class_Initialize()
    Set oFormulario = Nothing
    Set oTabela = Nothing
    Set oAutoIncremento = Nothing
    mCodigo = 0
    mIndex = 1
End Sub

Private Sub Class_Terminate()
    Set oFormulario = Nothing
    Set oTabela = Nothing
    mCodigo = 0
    mIndex = 1
End Sub

Private Function ObterProximoCodigo(nomeColuna As String) As Long
    With oAutoIncremento.ListColumns(nomeColuna).DataBodyRange(1, 1)
        ObterProximoCodigo = .Value
        .Value = .Value + 1
    End With
End Function

Private Function ObterIndiceRegistro() As Long
    Dim rg As Range
    Set rg = oTabela.ListColumns("Codigo").DataBodyRange

    On Error Resume Next
    ObterIndiceRegistro = 1
    ObterIndiceRegistro = Application.WorksheetFunction.Match(mCodigo, rg, 0)
    On Error GoTo 0

    Set rg = Nothing
End Function

Private Sub LimparCampos()
    Dim ctl As MSForms.Control
    
    For Each ctl In oFormulario.Controls
        Select Case TypeName(ctl)
            Case "TextBox", "ComboBox"
                ctl.Value = ""
        End Select
    Next ctl

    Set ctl = Nothing
End Sub

Private Sub DesbloquearCampos()
    Dim ctl As MSForms.Control
    
    For Each ctl In oFormulario.Controls
        ctl.Enabled = InStr(1, ctl.Tag, "#") = 0
    Next ctl

    Set ctl = Nothing
End Sub

Private Sub CarregarRegistro()
    Dim i As Long
    Dim nomeColuna As String

    With oTabela
        For i = 1 To .ListColumns.Count
            nomeColuna = .ListColumns(i).Name

            On Error Resume Next

            With oFormulario.Controls("tx" & nomeColuna)
                .Value = .ListColumns(i).DataBodyRange(mIndex, 1).Value
                .Enabled = False
            End With
            
            On Error GoTo 0
        Next i
    End With
End Sub

Private Sub ExibirNavegacao(exibir As Boolean)
    On Error Resume Next

    With oFormulario.Controls
        .Item("frameNavegacao").Visible = exibir
        .Item("frameCadastro").Visible = Not exibir
        .Item("frameAcao").Visible = exibir
        .Item("frameImagem").Visible = Not exibir
        .Item("txInfoAcao").Caption = ""
        .Item("btPesquisarCep").Visible = Not exibir
    End With
    
    On Error GoTo 0
End Sub

Private Sub GravarRegistro(nomeIndice As String)
    Dim campo As MSForms.TextBox
    Set campo = oFormulario.Controls(nomeIndice)
    
    mCodigo = campo.Value
    mIndex = ObterIndiceRegistro

    Dim i As Long
    Dim nomeColuna As String
    Dim valorCampo As Variant
    
    With oTabela
        If mCodigo = 0 Then
            .ListRows.Add
            mIndex = .ListRows.Count
            campo.Value = ObterProximoCodigo(oTabela.Name)
        End If

        For i = 1 To .ListColumns.Count
            nomeColuna = .ListColumns(i).Name
            valorCampo = oFormulario.Controls("tx" & nomeColuna).Value

            On Error Resume Next
            Debug.Print nomeColuna, valorCampo
            .ListColumns(i).DataBodyRange(mIndex, 1).Value = Trim(valorCampo)
            On Error GoTo 0
        Next i
    End With

    Set campo = Nothing

    ThisWorkbook.Save
End Sub

Private Function ExcluirRegistro(nomeIndice As String) As VbMsgBoxResult
    ExcluirRegistro = MsgBox("Você deseja realmente excluir este registro?", vbQuestion + vbYesNo)

    If ExcluirRegistro = vbNo Then Exit Function

    Dim campo As MSForms.TextBox
    Set campo = oFormulario.Controls(nomeIndice)

    mCodigo = campo.Value

    Dim indexRegistro As Long
    mIndex = ObterIndiceRegistro

    oTabela.ListRows(mIndex).Delete

    mCodigo = 0
    
    MsgBox "O registro foi excluído com sucesso", vbInformation
End Function

Public Sub CarregarImagem(campoImagem As MSForms.Image, campoCaminhoImagem As MSForms.TextBox, Optional Excluir As Boolean = False)
    Dim caminho As Variant

    Select Case True
        Case Excluir = True
            caminho = ""
            
        Case campoCaminhoImagem.Value <> ""
            caminho = campoCaminhoImagem.Value
            
        Case Else
            caminho = Application.GetOpenFilename(FileFilter : = "Imagens BMP (*.bmp), *.bmp", Title : = "Selecione uma imagem BMP")
            If caminho = False Then caminho = ThisWorkbook.Path & "\images\default.bmp"
    End Select

    campoCaminhoImagem.Value = caminho
    campoImagem.Picture = LoadPicture(caminho, 128, 180, Monochrome)
End Sub

Public Sub NavegarEntreRegistros(codigo As Long, Optional mover As String = "")
    Dim ultimaLinha As Long
    ultimaLinha = oTabela.ListRows.Count

    mCodigo = codigo

    Select Case mover
        Case "--"
            mIndex = 1

        Case "++"
            mIndex = ultimaLinha
        
        Case "+", "-"
            mIndex = ObterIndiceRegistro(mover)
    End Select

    Select Case True
        Case mIndex = 1 And mover = "-"
            MsgBox "Já está no primeiro registro", vbExclamation

        Case mIndex = ultimaLinha And mover = "+"
            MsgBox "Já está no último registro", vbExclamation

        Case Else
            CarregarRegistro
            ExibirNavegacao True
    End Select
End Sub

Public Sub Incluir()
    LimparCampos
    DesbloquearCampos
    ExibirNavegacao False
    oFormulario.Controls("txCodigo").Value = 0
    oFormulario.Controls("txInfoAcao").Caption = "Inserindo um novo registro"
End Sub

Public Sub Excluir()
    If ExcluirRegistro("txCodigo") = vbNo Then Exit Sub
    NavegarEntreRegistros 0
    ExibirNavegacao True
End Sub

Public Sub Editar()
    DesbloquearCampos
    ExibirNavegacao False
    oFormulario.Controls("txInfoAcao").Caption = "Editando o registro atual"
End Sub

Public Sub Cancelar()
    NavegarEntreRegistros 0
    ExibirNavegacao True
End Sub

Public Sub Salvar()
    GravarRegistro "txCodigo"
    ExibirNavegacao True
    NavegarEntreRegistros 0, "++"
End Sub