VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim oForm As MSForms.UserForm
Dim mIndex As Long
Dim mCodigo As Long

Private Function tbClientes() As ListObject
    Set tbClientes = shAuxBancoDados.ListObjects("tbClientes")
End Function

Private Function tbProximoId() As ListObject
    Set tbProximoId = shAuxBancoDados.ListObjects("tbProximoId")
End Function

Private Sub Class_Initialize()
    Set oForm = Nothing
    mCodigo = 0
End Sub

Private Sub Class_Terminate()
    Set oForm = Nothing
    mCodigo = 0
End Sub

Public Property Set Formulario(ByVal vNewValue As Object)
    Set oForm = vNewValue
End Property

Public Sub NavegarEntreRegistros(ByVal codigoCliente As Long, Optional mover As Long = 0)
    mCodigo = codigoCliente
    mIndex = 1
    
    If mover <> 9 Then
        mIndex = ObterClienteIndex(mover)
    End If
    
    CarregarRegistro
End Sub

Private Function ObterClienteIndex(Optional mover As Long = 0) As Long
    Dim lo As ListObject
    Set lo = tbClientes
    
    ObterClienteIndex = 0
    
    If mCodigo = 0 Then
        On Error Resume Next
        ObterClienteIndex = lo.ListRows.Count
        On Error GoTo 0
        Exit Function
    End If
    
    Dim rg As Range
    Set rg = lo.ListColumns("Codigo").DataBodyRange
    
    On Error Resume Next
    ObterClienteIndex = Application.WorksheetFunction.Match(mCodigo, rg, 0)
    On Error GoTo 0
    
    Select Case True
        Case ObterClienteIndex = 1 And mover = -1
            MsgBox "Já está no primeiro registro", vbExclamation
        
        Case ObterClienteIndex = lo.ListRows.Count And mover = 1
            MsgBox "Já está no último registro", vbExclamation
        
        Case Else
            ObterClienteIndex = ObterClienteIndex + mover

    End Select
    
    Set rg = Nothing
End Function


Public Sub Incluir()
    LimparCampos
    DesbloquearCampos
    ExibirNavegacao False
    RemoverImagem
    oForm.Controls("txCodigo").Value = 0
    oForm.Controls("txInfoAcao").Caption = "Inserindo um novo registro"
    oForm.Controls("btExcluirImg").Visible = False
    oForm.Controls("btEditarImg").Left = 42
End Sub

Public Sub Excluir()
    If ExcluirRegistro = vbNo Then Exit Sub
    NavegarEntreRegistros 0
    ExibirNavegacao True
End Sub

Public Sub Editar()
    DesbloquearCampos
    ExibirNavegacao False
    oForm.Controls("txInfoAcao").Caption = "Editando o registro atual"
    oForm.Controls("btExcluirImg").Visible = True
    oForm.Controls("btEditarImg").Left = 30
End Sub

Public Sub Cancelar()
    NavegarEntreRegistros 0
    ExibirNavegacao True
End Sub

Public Sub Salvar()
    GravarRegistro
    NavegarEntreRegistros 0
    ExibirNavegacao True
End Sub

Private Sub CarregarRegistro()
    Dim lo As ListObject
    Set lo = tbClientes
    
    Dim i As Long
    Dim nomeColuna As String
    Dim ctl As Object
    
    
    For i = 1 To lo.ListColumns.Count
        nomeColuna = lo.ListColumns(i).Name
        On Error Resume Next
        
        With oForm.Controls("tx" & nomeColuna)
            .Value = lo.ListColumns(i).DataBodyRange(mIndex, 1).Value
            .Enabled = False
        End With
        
        On Error GoTo 0
    Next i
    
    CarregarImagem lo.ListColumns("Imagem").DataBodyRange(mIndex, 1).Value
    ExibirNavegacao True
End Sub

Private Sub LimparCampos()
    Dim ctl As MSForms.Control
    For Each ctl In oForm.Controls
        Select Case TypeName(ctl)
            Case "TextBox", "ComboBox"
                ctl.Value = ""
        End Select
    Next ctl
    
    Set ctl = Nothing
End Sub

Private Sub DesbloquearCampos()
    Dim ctl As MSForms.Control
    For Each ctl In oForm.Controls
        ctl.Enabled = InStr(1, ctl.Tag, "#") = 0
    Next ctl
    Set ctl = Nothing
End Sub

Private Function ObterProximoCodigo() As Long
    With tbProximoId.ListColumns("ProximoID").DataBodyRange(1, 1)
        ObterProximoCodigo = .Value
        .Value = .Value + 1
    End With
End Function

Private Sub ExibirNavegacao(ByVal exibirControle As Boolean)
    oForm.Controls("frameNavegacao").Visible = exibirControle
    oForm.Controls("frameCadastro").Visible = Not exibirControle
    oForm.Controls("frameAcao").Visible = exibirControle
    oForm.Controls("frameImagem").Visible = Not exibirControle
    oForm.Controls("txInfoAcao").Caption = ""
    oForm.Controls("btPesquisarCep").Visible = Not exibirControle
End Sub

Private Sub GravarRegistro()
    Dim lo As ListObject
    Set lo = tbClientes
    
    mCodigo = oForm.Controls("txCodigo").Value
    
    Dim indexRegistro As Long
    mIndex = ObterClienteIndex

    If mCodigo = 0 Then
        lo.ListRows.Add
        mIndex = mIndex + 1
    End If
    
    Dim i As Long
    Dim nomeColuna As String
    
    For i = 1 To lo.ListColumns.Count
        nomeColuna = lo.ListColumns(i).Name
        On Error Resume Next
        With oForm.Controls("tx" & nomeColuna)
            lo.ListColumns(i).DataBodyRange(mIndex, 1).Value = Trim(.Value)
        End With
        On Error GoTo 0
    Next i
    
    If mCodigo = 0 Then
        lo.ListColumns("Codigo").DataBodyRange(mIndex, 1).Value = ObterProximoCodigo
    End If
    
    ThisWorkbook.Save
End Sub

Private Function ExcluirRegistro() As VbMsgBoxResult
    ExcluirRegistro = MsgBox("Você deseja realmente excluir este registro?", vbQuestion + vbYesNo)
    
    If ExcluirRegistro = vbNo Then Exit Function
    
    Dim lo As ListObject
    Set lo = tbClientes
    
    mCodigo = oForm.Controls("txCodigo").Value
    
    Dim indexRegistro As Long
    mIndex = ObterClienteIndex
    
    lo.ListRows(mIndex).Delete
    
    MsgBox "O registro foi excluído com sucesso", vbInformation
End Function

Private Sub CarregarImagem(ByVal caminho As String)
    If caminho = "" Then
        caminho = ThisWorkbook.Path & "\default.bmp"
    End If
    
    oForm.Controls("imgCliente").Picture = LoadPicture(caminho, 128, 228, Monochrome)
End Sub

Public Sub SelecionarImagem()
    Dim caminhoImagem As Variant

    caminhoImagem = Application.GetOpenFilename( _
        FileFilter:="Imagens BMP (*.bmp), *.bmp", _
        Title:="Selecione uma imagem BMP")

    If caminhoImagem <> False Then
        oForm.Controls("txImagem").Value = caminhoImagem
        CarregarImagem caminhoImagem
    End If
End Sub

Public Sub RemoverImagem()
    oForm.Controls("txImagem").Value = ""
    CarregarImagem ""
End Sub


