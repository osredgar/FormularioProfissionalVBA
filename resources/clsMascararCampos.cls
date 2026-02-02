VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMascararCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'KEYDOWN: ocorra quando o usuário pressiona uma tecla
'KEYPRESS: ocorre quando o usuário pressiona e libera uma tecla

Public WithEvents mAno As MSForms.TextBox
Attribute mAno.VB_VarHelpID = -1
Public WithEvents mCep As MSForms.TextBox
Attribute mCep.VB_VarHelpID = -1
Public WithEvents mData As MSForms.TextBox
Attribute mData.VB_VarHelpID = -1
Public WithEvents mCnpj As MSForms.TextBox
Attribute mCnpj.VB_VarHelpID = -1
Public WithEvents mCpf As MSForms.TextBox
Attribute mCpf.VB_VarHelpID = -1
Public WithEvents mMaiusculo As MSForms.TextBox
Attribute mMaiusculo.VB_VarHelpID = -1
Public WithEvents mMinusculo As MSForms.TextBox
Attribute mMinusculo.VB_VarHelpID = -1
Public WithEvents mTelefone As MSForms.TextBox
Attribute mTelefone.VB_VarHelpID = -1
Public WithEvents mEmail As MSForms.TextBox
Attribute mEmail.VB_VarHelpID = -1
Public WithEvents mNumeroTexto As MSForms.TextBox
Attribute mNumeroTexto.VB_VarHelpID = -1

Public WithEvents Decimal2 As MSForms.TextBox
Attribute Decimal2.VB_VarHelpID = -1
Public WithEvents Decimal4 As MSForms.TextBox
Attribute Decimal4.VB_VarHelpID = -1
Public WithEvents Decimal6 As MSForms.TextBox
Attribute Decimal6.VB_VarHelpID = -1
Public WithEvents Hora As MSForms.TextBox
Attribute Hora.VB_VarHelpID = -1
Public WithEvents Inteiro As MSForms.TextBox
Attribute Inteiro.VB_VarHelpID = -1
Public WithEvents Letra As MSForms.TextBox
Attribute Letra.VB_VarHelpID = -1


Private Function ObterApenasNumeros(ByVal pTexto As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Global = True
        .Pattern = "\D"
    End With
    
    ObterApenasNumeros = re.Replace(Trim(pTexto), "")

    Set re = Nothing
End Function

Public Function ValidarEmail() As Boolean
    ValidarEmail = True
    
    If mEmail.Value = "" Then Exit Function
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Global = True
        .MultiLine = True
        .Pattern = "\S+@\S+\.\S+"
    End With
    
    Dim texto As String
    texto = mEmail.Value
    
    ValidarEmail = re.test(texto)

    If Not ValidarEmail Then
        MsgBox "E-mail em formato inválido", vbExclamation
    End If

    Set re = Nothing
End Function


Public Function ValidarTelefone() As Boolean
    ValidarTelefone = True
    
    If mTelefone.Value = "" Then Exit Function
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Global = True
        .MultiLine = True
        .Pattern = "^\(?\d{2}\)?\s?(9?\d{4})-?\d{4}$"
    End With
    
    Dim texto As String
    texto = mTelefone.Value
    
    ValidarTelefone = re.test(texto)
    
    texto = ObterApenasNumeros(texto)
        
    If Not ValidarTelefone Then
        MsgBox "Formato de telefone inválido", vbExclamation
    End If
    
    Set re = Nothing
End Function

Private Function RemoverAcentos(ByVal texto As String, Optional tipoTexto As VbStrConv = vbUpperCase) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
        
    With re
        .Global = True
        .IgnoreCase = True
        .Pattern = "[áàãâä]"
        texto = .Replace(texto, "a")
        .Pattern = "[éèêë]"
        texto = .Replace(texto, "e")
        .Pattern = "[íìîï]"
        texto = .Replace(texto, "i")
        .Pattern = "[óòõôö]"
        texto = .Replace(texto, "o")
        .Pattern = "[úùûü]"
        texto = .Replace(texto, "u")
        .Pattern = "[ç]"
        texto = .Replace(texto, "c")
        .Pattern = "[ñ]"
        texto = .Replace(texto, "n")
    End With
    
    RemoverAcentos = StrConv(texto, tipoTexto)

    Set re = Nothing
End Function


Private Sub mAno_KeyPress(ByVal KeyAscii As ReturnInteger)
    mAno.MaxLength = 4
    
    Select Case KeyAscii
        Case 9, 13, 48 To 57
            KeyAscii = KeyAscii
        
        Case Else
            KeyAscii = 0
    End Select
End Sub

Public Function ValidarAno() As Boolean
    With mAno
        Dim comprimento As Long
        comprimento = .MaxLength
        
        Dim texto As String
        texto = .Value
        
        Dim anoMax As Long
        anoMax = Year(Date) + 50
    End With
    
    Select Case True
        Case texto = ""
            ValidarAno = True
        
        Case Trim(Len(texto)) <> comprimento
            ValidarAno = False
            MsgBox "Informe o ano com " & comprimento & " dígitos", vbExclamation

        Case CInt(texto) < 1900 Or CInt(texto) > anoMax
            ValidarAno = False
            MsgBox "Digite um ano entre 1900 e " & anoMax, vbExclamation
        
        Case Else
            ValidarAno = True
    End Select
End Function

Private Sub mCep_KeyPress(ByVal KeyAscii As ReturnInteger)
    Select Case KeyAscii
        Case 9, 13, 48 To 57
            KeyAscii = KeyAscii
        
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
    
    With mCep
        .MaxLength = 9
        If .SelStart = 5 Then .SelText = "-"
    End With
End Sub

Public Function ValidarCep() As Boolean
    ValidarCep = True
    
    If mCep.Value = "" Then Exit Function
       
    Dim numeros As String
    numeros = ObterApenasNumeros(mCep.Value)
    
    Select Case True
        Case Len(numeros) <> 8
            ValidarCep = False
            
        Case numeros = String(8, Left(numeros, 1))
            ValidarCep = False
    End Select

    If Not ValidarCep Then
        MsgBox "CEP inválido", vbExclamation
    End If
End Function

Private Sub mCnpj_KeyPress(ByVal KeyAscii As ReturnInteger)
    Select Case KeyAscii
        Case 9, 13, 48 To 57
            KeyAscii = KeyAscii
        
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
    
    With mCnpj
        .MaxLength = 18
        If .SelStart = 2 Then .SelText = "."
        If .SelStart = 6 Then .SelText = "."
        If .SelStart = 10 Then .SelText = "/"
        If .SelStart = 15 Then .SelText = "-"
    End With
End Sub

Public Function ValidarCnpj() As Boolean
    ValidarCnpj = True
    
    If mCnpj.Value = "" Then Exit Function
    
    Dim numeros As String
    numeros = ObterApenasNumeros(mCnpj.Value)

    If Len(numeros) <> 14 Then Exit Function
    
    Dim pesos1 As Variant
    pesos1 = Array(5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    
    Dim pesos2 As Variant
    pesos2 = Array(6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    
    Dim soma As Long
    soma = 0
    
    Dim i As Long
    For i = 0 To 11
        soma = soma + CLng(Mid(numeros, i + 1, 1)) * pesos1(i)
    Next i
    
    Dim resto As Long
    resto = soma Mod 11
    
    Dim verificador As String
    verificador = IIf(resto < 2, 0, 11 - resto)

    soma = 0
    
    For i = 0 To 12
        soma = soma + CLng(Mid(numeros, i + 1, 1)) * pesos2(i)
    Next i

    resto = soma Mod 11
    
    verificador = verificador & IIf(resto < 2, 0, 11 - resto)

    ValidarCnpj = (verificador = Right$(numeros, 2))
                   
    If Not ValidarCnpj Then
        MsgBox "CNPJ inválido", vbExclamation
    End If
End Function


Private Sub mCpf_KeyPress(ByVal KeyAscii As ReturnInteger)
    Select Case KeyAscii
        Case 9, 13, 48 To 57
            KeyAscii = KeyAscii
        
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
        
    With mCpf
        .MaxLength = 14
        If .SelStart = 3 Then .SelText = "."
        If .SelStart = 7 Then .SelText = "."
        If .SelStart = 11 Then .SelText = "-"
    End With
End Sub

Public Function ValidarCpf() As Boolean
    ValidarCpf = True
    
    If mCpf.Value = "" Then Exit Function
    
    Dim numeros As String
    numeros = ObterApenasNumeros(mCpf.Value)

    If Len(numeros) <> 11 Then Exit Function

    If numeros = String(11, Left(numeros, 1)) Then Exit Function

    Dim soma As Long
    soma = 0
    
    Dim i As Long
    For i = 1 To 9
        soma = soma + CLng(Mid(numeros, i, 1)) * (11 - i)
    Next i
    
    Dim resto As Long
    resto = (soma * 10) Mod 11
    
    Dim verificador As String
    verificador = IIf(resto = 10, 0, resto)

    soma = 0
    
    For i = 1 To 10
        soma = soma + CLng(Mid(numeros, i, 1)) * (12 - i)
    Next i

    resto = (soma * 10) Mod 11
    
    verificador = verificador & IIf(resto = 10, 0, resto)

    ValidarCpf = (verificador = Right$(numeros, 2))
                  
    If Not ValidarCpf Then
        MsgBox "CPF inválido", vbExclamation
    End If
End Function


Private Sub mData_KeyPress(ByVal KeyAscii As ReturnInteger)
    Select Case KeyAscii
        Case 9, 13, 48 To 57
            KeyAscii = KeyAscii
        
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
    
    With mData
        .MaxLength = 10
        If .SelStart = 2 Then .SelText = "/"
        If .SelStart = 5 Then .SelText = "/"
    End With
End Sub

Public Function ValidarData() As Boolean
    Dim texto As String
    texto = Trim(mData.Value)
    
    Dim anoMax As Long
    anoMax = Year(Date) + 50
    
    Select Case True
        Case texto = ""
            ValidarData = True
        
        Case Not IsDate(texto)
            MsgBox "Data inválida", vbExclamation
            ValidarData = False

        Case Year(CDate(texto)) < 1900 Or Year(CDate(texto)) > anoMax
            MsgBox "Digite uma data entre 01/01/1900 e 31/12/" & anoMax, vbExclamation
            ValidarData = False
        
        Case Else
            mData.Value = FormatDateTime(texto, vbShortDate)
            ValidarData = True
    End Select
End Function


Private Sub Decimal2_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    Select Case (KeyCode)
        Case 9, 13
            Select Case Decimal2
                Case ""
                    Let Decimal2 = Format(0, "#,##0.00")
                Case Else
                    Let Decimal2 = Format(Decimal2, "#,##0.00")
            End Select
        Case Else
            Exit Sub
    End Select
End Sub


Private Sub Decimal2_KeyPress(ByVal KeyAscii As ReturnInteger)
    Let Decimal2.TextAlign = 3
    Let Decimal2.MaxLength = 13

    Dim v As Integer
    Let v = InStr(1, Decimal2.Value, ",", vbTextCompare)

    Select Case (KeyAscii)
        Case 9, 13, 45, 48 To 57
            KeyAscii = KeyAscii
        Case 44
            KeyAscii = KeyAscii
            If (v <> 0) Then KeyAscii = 0
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
    
    Select Case Decimal2.SelStart
        Case Decimal2.MaxLength - 3
            Decimal2.SelText = ","
        Case Else
            If (v <> 0) Then Decimal2.MaxLength = v + 2
    End Select
End Sub


Private Sub Decimal4_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    Select Case (KeyCode)
        Case 9, 13
            Select Case Decimal4
                Case ""
                    Let Decimal4 = FormatNumber(0, 4, -1, 0, -1)
                Case Else
                    Let Decimal4 = FormatNumber(Decimal4, 4, -1, 0, -1)
            End Select
        Case Else
            Exit Sub
    End Select
End Sub


Private Sub Decimal4_KeyPress(ByVal KeyAscii As ReturnInteger)
    Let Decimal4.TextAlign = 3
    Let Decimal4.MaxLength = 15
    
    Select Case (KeyAscii)
        Case 9, 13, 44, 45, 48 To 57
            Let KeyAscii = KeyAscii
        Case Else
            Let KeyAscii = 0
            Exit Sub
    End Select
    
    Select Case Decimal4.SelStart
        Case Decimal4.MaxLength - 5
            Let Decimal4.SelText = ","
        Case Else
            Dim v As Integer
            Let v = InStr(1, Decimal4.Value, ",", vbTextCompare)
            If (v <> 0) Then Let Decimal4.MaxLength = v + 4
    End Select
End Sub


Private Sub Decimal6_KeyPress(ByVal KeyAscii As ReturnInteger)
    Let Decimal4.TextAlign = 3
    Let Decimal4.MaxLength = 17
    
    Select Case (KeyAscii)
        Case 9, 13, 44, 45, 48 To 57
            Let KeyAscii = KeyAscii
        Case Else
            Let KeyAscii = 0
            Exit Sub
    End Select
    
    Select Case Decimal4.SelStart
        Case Decimal4.MaxLength - 7
            Let Decimal4.SelText = ","
        Case Else
            Dim v As Integer
            Let v = InStr(1, Decimal4.Value, ",", vbTextCompare)
            If (v <> 0) Then Let Decimal4.MaxLength = v + 6
    End Select
End Sub


Private Sub Decimal6_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    Select Case (KeyCode)
        Case 9, 13
            Select Case Decimal6
                Case ""
                    Let Decimal6 = FormatNumber(0, 6, -1, 0, -1)
                Case Else
                    Let Decimal6 = FormatNumber(Decimal6, 6, -1, 0, -1)
            End Select
        Case Else
            Exit Sub
    End Select
End Sub


Private Sub Hora_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 9 Or KeyCode = 13 Then
        
        If Hora.Value = "" Then
            Exit Sub
        ElseIf UBound(Split(Hora.Value, ":")) = 0 Then
            KeyCode = 0
            MsgBox "Hora em formato inválido", vbExclamation
        ElseIf Split(Hora.Value, ":")(0) > 23 Or Split(Hora.Value, ":")(1) > 59 Then
            KeyCode = 0
            MsgBox "Hora em formato inválido", vbExclamation
            Hora.Value = ""
        Else
            Hora.Value = Format(Hora.Value, "Short Time")
        End If
        
    End If
    
End Sub


Private Sub Hora_KeyPress(ByVal KeyAscii As ReturnInteger)
    
    Hora.MaxLength = 5
    
    If (KeyAscii < 48 Or KeyAscii > 57) _
    And (KeyAscii <> 13 Or KeyAscii <> 9) Then
        KeyAscii = 0
    ElseIf Hora.SelStart = 2 Then
        Hora.SelText = ":"
    End If
    
End Sub


Private Sub Inteiro_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    Select Case (KeyCode)
        Case 9, 13
            Select Case IsNull(Inteiro)
                Case True
                    Let Inteiro = Null
                Case Else
                    Let Inteiro = FormatNumber(Val(Inteiro), 0, -1, 0, -1)
            End Select
        Case Else
            Exit Sub
    End Select
End Sub


Private Sub Inteiro_KeyPress(ByVal KeyAscii As ReturnInteger)
    Inteiro.TextAlign = fmTextAlignRight
    Select Case KeyAscii
        Case vbKeyTab, vbKeyReturn, vbKeySubtract, vbKeyBack, vbKeyDelete, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub Letra_KeyPress(ByVal KeyAscii As ReturnInteger)
    Select Case (KeyAscii)
        Case 65 To 90, 97 To 122
            Let KeyAscii = KeyAscii
        Case Else
            Let KeyAscii = 0
    End Select
End Sub

Private Sub mNumeroTexto_KeyPress(ByVal KeyAscii As ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub mMaiusculo_Change()
    Dim texto As String
    texto = mMaiusculo.Value
    mMaiusculo.Value = RemoverAcentos(texto, vbUpperCase)
End Sub

Private Sub mMaiusculo_KeyPress(ByVal KeyAscii As ReturnInteger)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub mMinusculo_Change()
    Dim texto As String
    texto = mMinusculo.Value
    mMinusculo.Value = RemoverAcentos(texto, vbLowerCase)
End Sub

Private Sub mMinusculo_KeyPress(ByVal KeyAscii As ReturnInteger)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

Private Sub mTelefone_KeyPress(ByVal KeyAscii As ReturnInteger)
    Select Case KeyAscii
        Case 9, 13, 48 To 57
            KeyAscii = KeyAscii
        
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
    
    Dim inicioNumero As Long
    inicioNumero = 9
    
    On Error Resume Next
    inicioNumero = IIf(Mid$(mTelefone.Value, 6, 1) = 9, 10, 9)
    On Error GoTo 0
    
    With mTelefone
        .MaxLength = IIf(inicioNumero = 9, 14, 15)
        If .SelStart = 0 Then .SelText = "("
        If .SelStart = 3 Then .SelText = ") "
        If .SelStart = inicioNumero Then .SelText = "-"
    End With
End Sub


Private Sub mEmail_KeyPress(ByVal KeyAscii As ReturnInteger)
    KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
    
    With mEmail
        .MaxLength = 255
    End With
End Sub



