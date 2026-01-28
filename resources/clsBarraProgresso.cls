VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBarraProgresso"
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

Private tamMax As Integer
Private fbarra As Object
Private txtPorc As String

Private Sub Class_Initialize()
    Set fbarra = New frmBarraProgresso
    
    With fbarra
        .Show vbModeless
        tamMax = .txbarra.Width
        .txbarra.Width = 0
        .Caption = "Concluído 0.0%"
    End With
End Sub

Private Sub Class_Terminate()
    Unload fbarra
    Set fbarra = Nothing
End Sub

Public Sub Atualizar(ByVal porc As Double, Optional processo As String = "")
    On Error GoTo TratarErro
    
    With fbarra
        .txbarra.Width = tamMax * porc
        
        txtPorc = "Concluído " & FormatPercent(porc, 2, vbTrue, vbFalse, vbTrue)
        
        If processo = "" Then
            .Caption = txtPorc
        Else
            .Caption = StrConv(Trim(processo), 3) & " - " & txtPorc
        End If
    End With
    
    DoEvents
    
Finalizar:
    On Error GoTo 0
    Exit Sub

TratarErro:
    Debug.Print "clsBarraProgresso > Atualizar" & vbNewLine & "Erro: " & Abs(Err.Number) & vbNewLine & "Descricao: " & Err.Description
    Resume Finalizar
End Sub
