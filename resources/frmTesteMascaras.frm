VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTesteMascaras 
   Caption         =   "Máscaras de campos"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14340
   OleObjectBlob   =   "frmTesteMascaras.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTesteMascaras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oMask() As New clsMascararCampos

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(1).ValidarAno
End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(2).ValidarCep
End Sub

Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(3).ValidarData
    If Cancel Then TextBox3.Value = ""
End Sub

Private Sub TextBox6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(6).ValidarCnpj
End Sub

Private Sub TextBox7_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(7).ValidarCpf
End Sub

Private Sub TextBox8_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(8).ValidarTelefone
End Sub

Private Sub TextBox9_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not oMask(9).ValidarEmail
End Sub

Private Sub UserForm_Activate()
    ReDim Preserve oMask(1 To 9)
    Set oMask(1).mAno = TextBox1
    Set oMask(2).mCep = TextBox2
    Set oMask(3).mData = TextBox3
    Set oMask(4).mMaiusculo = TextBox4
    Set oMask(5).mMinusculo = TextBox5
    Set oMask(6).mCnpj = TextBox6
    Set oMask(7).mCpf = TextBox7
    Set oMask(8).mTelefone = TextBox8
    Set oMask(9).mEmail = TextBox9
End Sub
