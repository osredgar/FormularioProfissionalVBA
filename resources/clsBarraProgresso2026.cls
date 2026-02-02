VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBarraProgresso2026"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**********************************************************************
'ESTA CLASSE SERVE PARA TRATATIVAS DA BARRA DE PROGRESSO
'DEVE-SE IMPORTAR O FORMULARIO DA BARRA DE PROGRESSO
'ESTA CLASSE DEVE SER INSTANCIADA DENTRO DO FORMULARIO OU
'PLANILHA QUE DESEJA UTILIZAR
'**********************************************************************

Private mTamanhoMaximo As Integer
Private oBarra As frmBarraProgresso

Private Sub Class_Initialize()
    Set oBarra = New frmBarraProgresso
    
    mTamanhoMaximo = oBarra.barraProgresso.Width
    With oBarra
        .Show vbModeless
        .barraProgresso.Width = 0
        .infoBarra = "CONCLUÍDO 0.00%"
    End With
End Sub

Private Sub Class_Terminate()
    Unload oBarra
    Set oBarra = Nothing
End Sub

Public Sub Atualizar(ByVal porcentagem As Double, Optional nomeProcesso As String = "")
    On Error Resume Next
    
    Dim infoProgresso As String
    infoProgresso = "[processo] - CONCLUÍDO: [porcentagem]"

    With oBarra
        .barraProgresso.Width = mTamanhoMaximo * porcentagem
        
        infoProgresso = Replace(infoProgresso, "[porcentagem]", Format(porcentagem, "0.00%"))
        
        Select Case True
            Case Len(Trim(nomeProcesso)) <> 0
                infoProgresso = Replace(infoProgresso, "[processo]", UCase(nomeProcesso))
                
            Case Else
                infoProgresso = Replace(infoProgresso, "[processo] - ", "")
        End Select

        .infoBarra = infoProgresso
    End With
    
    DoEvents

    On Error GoTo 0
End Sub
