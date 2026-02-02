VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormularios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim oFormulario As MSForms.UserForm
Dim oTabela As ListObject
Dim oAutoIncremento As ListObject
Dim mIndice As Long
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
    mIndice = 1
End Sub

Private Sub Class_Terminate()
    Set oFormulario = Nothing
    Set oTabela = Nothing
    mCodigo = 0
    mIndice = 1
End Sub

Private Function ObterProximoCodigo() As Long
    With oAutoIncremento.Range(2, 1)
        ObterProximoCodigo = .Value
        .Value = .Value + 1
    End With
End Function

Private Function ObterIndiceRegistro(Optional mover As String = "") As Long
    Dim rg As Range
    Set rg = oTabela.ListColumns("Codigo").DataBodyRange
    
    On Error Resume Next
    ObterIndiceRegistro = 1
    ObterIndiceRegistro = Application.WorksheetFunction.Match(mCodigo, rg, 0)
    On Error GoTo 0
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
    Dim campo As Object
    
    With oTabela
        For i = 1 To .ListColumns.Count
            nomeColuna = .ListColumns(i).Name
            
            On Error Resume Next
            
            Set campo = oFormulario.Controls("tx" & nomeColuna)
            campo.Value = .ListColumns(i).DataBodyRange(mIndice, 1).Value
            campo.Enabled = False

            On Error GoTo 0
        Next i
    End With
    
    Set campo = Nothing
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
    mIndice = ObterIndiceRegistro
    
    Dim i As Long
    Dim nomeColuna As String
    Dim valorCampo As Variant

    With oTabela
        If mCodigo = 0 Then
            .ListRows.Add
            mIndice = .ListRows.Count
            campo.Value = ObterProximoCodigo
        End If
        
        For i = 1 To .ListColumns.Count
            nomeColuna = .ListColumns(i).Name
            valorCampo = oFormulario.Controls("tx" & nomeColuna).Value
            
            On Error Resume Next
            Debug.Print nomeColuna, valorCampo
            .ListColumns(i).DataBodyRange(mIndice, 1).Value = Trim(valorCampo)
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
    mIndice = ObterIndiceRegistro
    
    oTabela.ListRows(mIndice).Delete
    
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
            caminho = Application.GetOpenFilename(FileFilter:="Imagens BMP (*.bmp), *.bmp", Title:="Selecione uma imagem BMP")
            If caminho = False Then caminho = ThisWorkbook.Path & "\images\default.bmp"
    End Select
    
    campoCaminhoImagem.Value = caminho
    campoImagem.Picture = LoadPicture(caminho, 128, 180, Monochrome)
End Sub

Public Sub NavegarEntreRegistros(Optional mover As String = "")
    mCodigo = oFormulario.Controls("txCodigo").Value
    
    Dim ultimaLinha As Long
    ultimaLinha = oTabela.ListRows.Count
    
    mIndice = ObterIndiceRegistro(mover)

    Select Case True
        Case mover = "-" And mIndice = 1
            MsgBox "Já está no primeiro registro", vbExclamation
            Exit Sub
            
        Case mover = "+" And mIndice = ultimaLinha
            MsgBox "Já está no último registro", vbExclamation
            Exit Sub

        Case mover = "+"
            mIndice = mIndice + 1

        Case mover = "-"
            mIndice = mIndice - 1

        Case mover = "--"
            mIndice = 1

        Case mover = "++"
            mIndice = ultimaLinha
    End Select
    
    CarregarRegistro
    ExibirNavegacao True
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
    NavegarEntreRegistros
    ExibirNavegacao True
End Sub

Public Sub Editar()
    DesbloquearCampos
    ExibirNavegacao False
    oFormulario.Controls("txInfoAcao").Caption = "Editando o registro atual"
End Sub

Public Sub Cancelar()
    NavegarEntreRegistros
    ExibirNavegacao True
End Sub

Public Sub Salvar()
    GravarRegistro "txCodigo"
    ExibirNavegacao True
    NavegarEntreRegistros "++"
End Sub
