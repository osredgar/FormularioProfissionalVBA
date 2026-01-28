VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadClientes 
   Caption         =   "Clientes"
   ClientHeight    =   9615.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "cadClientes.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cadClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oCls As clsClientes
Dim oMask() As New clsMascararCampos
Dim oViaCep As New clsViaCep

Private Sub btCancelar_Click()
    oCls.Cancelar
End Sub

Private Sub btEditar_Click()
    oCls.Editar
End Sub

Private Sub btEditarImg_Click()
    oCls.SelecionarImagem
End Sub

Private Sub btExcluir_Click()
    oCls.Excluir
End Sub

Private Sub btExcluirImg_Click()
    oCls.RemoverImagem
End Sub

Private Sub btIncluir_Click()
    oCls.Incluir
End Sub

Private Sub btNavAnterior_Click()
    oCls.NavegarEntreRegistros txCodigo.Value, -1
End Sub

Private Sub btNavPrimeiro_Click()
    oCls.NavegarEntreRegistros 0, 9
End Sub

Private Sub btNavProximo_Click()
    oCls.NavegarEntreRegistros txCodigo.Value, 1
End Sub

Private Sub btNavUltimo_Click()
    oCls.NavegarEntreRegistros 0, 0
End Sub

Private Sub btPesquisarCep_Click()
    oViaCep.PreencherCepAutomaticamente Me
    Set oViaCep = Nothing
End Sub

Private Sub btSalvar_Click()
    oCls.Salvar
End Sub

Private Sub txCep_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(3).ValidarCep
End Sub

Private Sub txDataNasc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(1).ValidarData
    If Cancel Then txDataNasc.Value = ""
End Sub

Private Sub txEmail_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(4).ValidarEmail
End Sub

Private Sub txTelefone_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(2).ValidarTelefone
End Sub

Private Sub UserForm_Initialize()
    Set oCls = New clsClientes
    Set oCls.Formulario = Me
    btNavUltimo_Click
    AplicarMascaraCampos
End Sub

Private Sub UserForm_Terminate()
    Set oCls = Nothing
    
'    With ThisWorkbook
'        .Save
'        .Close
'    End With
End Sub


Public Sub AplicarMascaraCampos()
    ReDim Preserve oMask(1 To 9)
    Set oMask(1).mData = txDataNasc
    Set oMask(2).mTelefone = txTelefone
    Set oMask(3).mCep = txCep
    Set oMask(4).mEmail = txEmail
    Set oMask(5).mMaiusculo = txNome
    Set oMask(6).mMaiusculo = txEndereco
    Set oMask(7).mMaiusculo = txBairro
    Set oMask(8).mMaiusculo = txCidade
    Set oMask(9).mMaiusculo = txEstado
End Sub
