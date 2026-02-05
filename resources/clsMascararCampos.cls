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

Private mAtualizando As Boolean

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
Public WithEvents mDecimal As MSForms.TextBox
Attribute mDecimal.VB_VarHelpID = -1
Public WithEvents mHora As MSForms.TextBox
Attribute mHora.VB_VarHelpID = -1
Public WithEvents mInteiro As MSForms.TextBox
Attribute mInteiro.VB_VarHelpID = -1
Public WithEvents mLetra As MSForms.TextBox
Attribute mLetra.VB_VarHelpID = -1


Private Function ObterApenasNumeros(ByVal texto As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Global = True
        .Pattern = "\D"
        texto = .Replace(texto, "")
    End With
    
    ObterApenasNumeros = Trim(texto)

    Set re = Nothing
End Function

Private Function ObterApenasLetras(ByVal texto As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    texto = Trim(texto)
    
    With re
        .Global = True
        .Pattern = "[^A-Za-z¡¿¬√ƒ…» ÀÕÃŒœ”“‘’÷⁄Ÿ€‹«·‡‚„‰ÈËÍÎÌÏÓÔÛÚÙıˆ˙˘˚¸Á ]+"
        texto = .Replace(texto, "")
        .Pattern = " {2,}"
        texto = .Replace(texto, " ")
    End With
    
    ObterApenasLetras = texto

    Set re = Nothing
End Function

Private Function RemoverAcentos(ByVal texto As String, Optional tipoTexto As VbStrConv = vbProperCase) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    With re
        .Global = True
        .IgnoreCase = True
        .Pattern = "[·‡„‚‰Â]"
        texto = .Replace(texto, "a")
        .Pattern = "[¡¿√¬ƒ≈]"
        texto = .Replace(texto, "A")
        .Pattern = "[ÈËÍÎ]"
        texto = .Replace(texto, "e")
        .Pattern = "[…» À]"
        texto = .Replace(texto, "E")
        .Pattern = "[ÌÏÓÔ]"
        texto = .Replace(texto, "i")
        .Pattern = "[ÕÃŒœ]"
        texto = .Replace(texto, "I")
        .Pattern = "[ÛÚıÙˆ¯]"
        texto = .Replace(texto, "o")
        .Pattern = "[”“’‘÷ÿ]"
        texto = .Replace(texto, "O")
        .Pattern = "[˙˘˚¸]"
        texto = .Replace(texto, "u")
        .Pattern = "[⁄Ÿ€‹]"
        texto = .Replace(texto, "U")
        .Pattern = "[Á]"
        texto = .Replace(texto, "c")
        .Pattern = "[«]"
        texto = .Replace(texto, "C")
        .Pattern = "[Ò]"
        texto = .Replace(texto, "n")
        .Pattern = "[—]"
        texto = .Replace(texto, "N")
        .Pattern = "[Ê]"
        texto = .Replace(texto, "ae")
        .Pattern = "[∆]"
        texto = .Replace(texto, "AE")
        .Pattern = "[ú]"
        texto = .Replace(texto, "oe")
        .Pattern = "[å]"
        texto = .Replace(texto, "OE")
        .Pattern = "[]"
        texto = .Replace(texto, "d")
        .Pattern = "[–]"
        texto = .Replace(texto, "D")
        .Pattern = "[˛]"
        texto = .Replace(texto, "th")
        .Pattern = "[ﬁ]"
        texto = .Replace(texto, "TH")
        .Pattern = "[ﬂ]"
        texto = .Replace(texto, "ss")
        .Pattern = "[ﬂ]"
        texto = .Replace(texto, "SS")
        .Pattern = "[˝ˇy]"
        texto = .Replace(texto, "y")
        .Pattern = "[›üY]"
        texto = .Replace(texto, "Y")
        .Pattern = " {2,}"
        texto = .Replace(texto, " ")
    End With

    RemoverAcentos = StrConv(texto, tipoTexto)

    Set re = Nothing
End Function

Public Function EhAnoValido() As Boolean
    EhAnoValido = True
    
    Dim valor As String
    valor = mAno.Value
    
    Dim anoMax As Integer
    anoMax = Year(Date) + 50
    
    Select Case True
        Case Len(valor) = 0
            Exit Function
        
        Case Len(valor) < 4
            MsgBox "Informe o ano com 4 dÌgitos", vbExclamation
            EhAnoValido = False

        Case CInt(valor) < 1900 Or CInt(valor) > anoMax
            MsgBox "Digite um ano entre 1900 e " & anoMax, vbExclamation
            EhAnoValido = False
    End Select
End Function

Private Sub mAno_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyBack
            Exit Sub
            
        Case KeyCode = vbKeyV And Shift = 2
            Exit Sub
            
        Case KeyCode >= vbKey0 And KeyCode <= vbKey9, KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9
            Exit Sub
            
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub mAno_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With mAno
        .MaxLength = 4
    
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9, KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9
                Exit Sub
            
            Case Else
                KeyAscii = 0
        End Select
    End With
End Sub

Public Function EhCepValido() As Boolean
    EhCepValido = True
    
    Dim valor As String
    valor = ObterApenasNumeros(mCep.Value)
    
    Select Case True
        Case Len(Trim(valor)) = 0
            Exit Function

        Case Len(valor) <> 8
            MsgBox "CEP deve conter 8 dÌgitos", vbExclamation
            EhCepValido = False

        Case valor = String(8, Left(valor, 1))
            MsgBox "Digite um CEP v·lido", vbExclamation
            EhCepValido = False

        Case Else
            mCep.Value = CStr(Format(valor, "00000\-000"))
    End Select
End Function

Private Sub mCep_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyBack
            Exit Sub
            
        Case KeyCode = vbKeyV And Shift = 2
            Exit Sub
            
        Case KeyCode >= vbKey0 And KeyCode <= vbKey9, KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9
            Exit Sub
            
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub mCep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With mCep
        .MaxLength = 9
    
        Select Case True
            Case .SelStart = 5
                .SelText = "-"
            
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9, KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9
                Exit Sub
            
            Case Else
                KeyAscii = 0
        End Select
    End With
End Sub

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
        MsgBox "CNPJ inv·lido", vbExclamation
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
        MsgBox "CPF inv·lido", vbExclamation
    End If
End Function

Public Function EhDataValida() As Boolean
    EhDataValida = True
    
    Dim valor As String
    valor = Trim(mData.Value)
    
    Dim anoMax As Long
    anoMax = Year(Date) + 50
    
    Select Case True
        Case Len(valor) = 0
            Exit Function

        Case Len(valor) < 10
            MsgBox "A data deve conter 10 dÌgitos e estar no formato DD/MM/AAAA", vbExclamation
            EhDataValida = False

        Case Not IsDate(valor)
            MsgBox "Digite uma data v·lida", vbExclamation
            EhDataValida = False
        
        Case Right(valor, 4) < 1900 Or Right(valor, 4) > anoMax
            MsgBox "A data deve estar entre os anos 1900 e " & anoMax, vbExclamation
            EhDataValida = False
        
        Case Else
            mData.Value = Format$(valor, "dd/mm/yyyy")
    End Select
End Function

Private Sub mData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyBack
            Exit Sub
        
        Case KeyCode >= vbKey0 And KeyCode <= vbKey9, KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9
            Exit Sub
        
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub mData_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With mData
        .MaxLength = 10
        
        Select Case True
            Case .SelStart = 2, .SelStart = 5
                .SelText = "/"
            
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9, KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9
                Exit Sub
            
            Case Else
                KeyAscii = 0
        End Select
    End With
End Sub

Private Sub mDecimal_Change()
    If Len(Trim(mDecimal.Text)) = 0 Then Exit Sub
    
    If mAtualizando Then Exit Sub
    mAtualizando = True
    
    mDecimal.MaxLength = 20
    
    Dim valor As String
    valor = ObterApenasNumeros(mDecimal.Text)

    If Len(valor) > mDecimal.MaxLength Then
        valor = Left$(valor, mDecimal.MaxLength)
    End If
    
    mDecimal.Value = Format(CCur(valor / 100), "#,##0.00")

    mAtualizando = False
End Sub

Private Sub mDecimal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim valor As String
    
    Select Case True
        Case KeyCode = vbKeyBack, KeyCode = vbKeyReturn, KeyCode = vbKeyTab
            Exit Sub

        Case KeyCode = vbKeyV And Shift = 2
            Exit Sub

        Case KeyCode >= vbKey0 And KeyCode <= vbKey9, KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9
            On Error Resume Next
            valor = CDbl(mDecimal.Text) * 100
            If Len(valor) <= 15 Then valor = valor & CStr(KeyCode Mod 48)
            On Error GoTo 0
        
        Case Else
            KeyCode = 0
    End Select
    
    On Error Resume Next
    mDecimal.Value = Format(CCur(valor) / 100, "#,##0.00")
    On Error GoTo 0
    
    KeyCode = 0
End Sub

Private Sub mDecimal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    mDecimal.MaxLength = 20
    
    Select Case True
        Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Exit Sub
        
        Case Else
            KeyAscii = 0
    End Select
End Sub

Public Function EhHoraValida() As Boolean
    EhHoraValida = True
    
    Select Case True
        Case Len(Trim(mHora.Value)) = 0
            Exit Function

        Case Not IsDate(mHora.Value)
            MsgBox "Digite uma hora v·lida", vbExclamation
            EhHoraValida = False
        
        Case Else
            mHora.Value = CStr(Format(CDate(mHora.Value), "Short Time"))
    End Select
End Function

Private Sub mHora_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyBack
            Exit Sub
            
        Case KeyCode = vbKeyV And Shift = 2
            Exit Sub
            
        Case KeyCode >= vbKey0 And KeyCode <= vbKey9, KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9
            Exit Sub
            
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub mHora_KeyPress(ByVal KeyAscii As ReturnInteger)
    mHora.MaxLength = 5
    
    Select Case True
        Case mHora.SelStart = 2
            mHora.SelText = ":"
        
        Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9, KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9
            Exit Sub
        
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub mInteiro_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    Select Case (KeyCode)
        Case 9, 13
            Select Case IsNull(mInteiro)
                Case True
                    Let mInteiro = Null
                Case Else
                    Let mInteiro = FormatNumber(Val(mInteiro), 0, -1, 0, -1)
            End Select
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub mInteiro_KeyPress(ByVal KeyAscii As ReturnInteger)
    mInteiro.TextAlign = fmTextAlignRight
    Select Case KeyAscii
        Case vbKeyTab, vbKeyReturn, vbKeySubtract, vbKeyBack, vbKeyDelete, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub mLetra_Change()
    Dim texto As String
    texto = Trim(mLetra.Text)

    If Len(texto) = 0 Then Exit Sub

    If mAtualizando Then Exit Sub
    mAtualizando = True

    texto = ObterApenasLetras(texto)

    With mLetra
        .MaxLength = 255

        If Len(texto) > .MaxLength Then texto = Left$(texto, .MaxLength)

        .Value = Trim(texto)
    End With

    mAtualizando = False
End Sub

Private Sub mLetra_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyBack Or KeyCode = vbKeyDelete
            Exit Sub

        Case KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd
            Exit Sub

        Case KeyCode = vbKeyV And Shift = 2
            Exit Sub
            
        Case KeyCode >= 65 And KeyCode <= 90, KeyCode >= 97 And KeyCode <= 122, KeyCode = 32
            Exit Sub

        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub mLetra_KeyPress(ByVal KeyAscii As ReturnInteger)
    mLetra.MaxLength = 255
    
    Select Case True
        Case KeyAscii >= 65 And KeyAscii <= 90, KeyAscii >= 97 And KeyAscii <= 122, KeyAscii = 32
            Exit Sub

        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub mMaiusculo_Change()
    With mMaiusculo
        .MaxLength = 255
        .Value = RemoverAcentos(.Value, vbUpperCase)
    End With
End Sub

Private Sub mMaiusculo_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyBack Or KeyCode = vbKeyDelete
            Exit Sub

        Case KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd
            Exit Sub

        Case KeyCode = vbKeyV And Shift = 2
            Exit Sub

        Case KeyCode >= vbKey0 And KeyCode <= vbKey9
            Exit Sub

        Case KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9
            Exit Sub
            
        Case KeyCode >= vbKeyA And KeyCode <= vbKeyZ, KeyCode = vbKeySpace
            Exit Sub

        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub mMaiusculo_KeyPress(ByVal KeyAscii As ReturnInteger)
    mMaiusculo.MaxLength = 255
    KeyAscii = AscW(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub mMinusculo_Change()
    With mMinusculo
        .MaxLength = 255
        .Value = RemoverAcentos(.Value, vbLowerCase)
    End With
End Sub

Private Sub mMinusculo_KeyDown(ByVal KeyCode As ReturnInteger, ByVal Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyBack Or KeyCode = vbKeyDelete
            Exit Sub

        Case KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyHome Or KeyCode = vbKeyEnd
            Exit Sub

        Case KeyCode = vbKeyV And Shift = 2
            Exit Sub

        Case KeyCode >= vbKey0 And KeyCode <= vbKey9
            Exit Sub

        Case KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9
            Exit Sub
            
        Case KeyCode >= vbKeyA And KeyCode <= vbKeyZ, KeyCode = vbKeySpace
            Exit Sub

        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub mMinusculo_KeyPress(ByVal KeyAscii As ReturnInteger)
    mMinusculo.MaxLength = 255
    KeyAscii = AscW(LCase$(Chr$(KeyAscii)))
End Sub

Private Sub mNumeroTexto_Change()
    mNumeroTexto.Value = ObterApenasNumeros(mNumeroTexto.Value)
End Sub

Private Sub mNumeroTexto_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyBack Or KeyCode = vbKeyDelete
            Exit Sub
            
        Case KeyCode = vbKeyV And Shift = 2
            Exit Sub
            
        Case KeyCode >= vbKey0 And KeyCode <= vbKey9, KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9
            Exit Sub
            
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub mNumeroTexto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With mNumeroTexto
        .MaxLength = 255
    
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9, KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9
                Exit Sub
            
            Case Else
                KeyAscii = 0
        End Select
    End With
End Sub

Public Function EhEmailValido() As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Global = True
        .MultiLine = True
        .Pattern = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    End With
    
    Dim texto As String
    texto = Trim(mEmail.Value)
    
    Select Case True
        Case Len(texto) = 0
            EhEmailValido = True
        
        Case re.test(texto)
            EhEmailValido = True
        
        Case Else
            MsgBox "E-mail em formato inv·lido", vbExclamation
            EhEmailValido = False
    End Select

    Set re = Nothing
End Function

Private Sub mEmail_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyBack, vbKeyDelete, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd
            Exit Sub

        Case vbKeyV
            If Shift = 2 Then Exit Sub
    End Select
End Sub

Private Sub mEmail_KeyPress(ByVal KeyAscii As ReturnInteger)
    mEmail.MaxLength = 255
    KeyAscii = AscW(LCase$(Chr$(KeyAscii)))

    Select Case KeyAscii
        Case 48 To 57
            '0-9
            Exit Sub

        Case 65 To 90, 97 To 122
            'A-Z, a-z
            Exit Sub

        Case 45, 46, 64, 95
            '-.@_
            Exit Sub
    
        Case Else
            KeyAscii = 0
    End Select
End Sub

Public Function EhTelefoneValido() As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Global = True
        .MultiLine = True
        .Pattern = "^\(?\d{2}\)?\s?(9?\d{4})-?\d{4}$"
    End With
    
    Dim valor As String
    valor = Trim(mTelefone.Value)
    
    Select Case True
        Case Len(valor) = 0
            EhTelefoneValido = True
            
        Case re.test(valor)
            EhTelefoneValido = True
            
        Case Else
            MsgBox "Telefone digitado est· em um formato inv·lido", vbExclamation
            EhTelefoneValido = False
    End Select
    
    Set re = Nothing
End Function

Private Sub mTelefone_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyBack
            Exit Sub
            
        Case KeyCode = vbKeyV And Shift = 2
            Exit Sub
            
        Case KeyCode >= vbKey0 And KeyCode <= vbKey9, KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9
            Exit Sub
            
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub mTelefone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim inicioNumero As Integer
    inicioNumero = 9
    
    With mTelefone
        On Error Resume Next
        inicioNumero = IIf(Mid$(.Value, 6, 1) = 9, 10, 9)
        On Error GoTo 0
        
        .MaxLength = IIf(inicioNumero = 9, 14, 15)
        
        Select Case True
            Case .SelStart = 0
                .SelText = "("
            
            Case .SelStart = 3
                .SelText = ") "
                
            Case .SelStart = inicioNumero
                .SelText = "-"
            
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9, KeyAscii >= vbKeyNumpad0 And KeyAscii <= vbKeyNumpad9
                Exit Sub
            
            Case Else
                KeyAscii = 0
        End Select
    End With
End Sub
