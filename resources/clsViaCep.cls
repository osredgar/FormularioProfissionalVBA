VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsViaCep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim oReq As Object
Dim xmlDoc As Object

'====================================================================================================
'Desenvolvido por Edgar Santa Rosa (C) 2025
'FUNCIONALIDADE: CONSULTAR CEP ATRAVES DA API VIA CEP - "https://viaCep.com.br/ws/[CEP]/json/"
'Parametros: Formulário que será preenchido com os dados consultados e que contenha o campo de Cep.
'RETORNO: API retorna os dados em XML que será utilizado para preencher o formulário
'OBSERVACAO: API possui limitação de no máximo 3 consultas por minuto
'====================================================================================================

Private Sub Class_Initialize()
    On Error Resume Next
    Set oReq = CreateObject("MSXML2.XMLHTTP60")
    If oReq Is Nothing Then
        Set oReq = CreateObject("MSXML2.XMLHTTP")
    End If
    On Error GoTo 0

    On Error Resume Next
    Set xmlDoc = CreateObject("MSXML2.DOMDocument60")
    If xmlDoc Is Nothing Then
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    End If
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Set oReq = Nothing
    Set xmlDoc = Nothing
End Sub

Private Function EhCepValido(ByVal cep As String) As String
    EhCepValido = ""
    
    If cep = "" Then Exit Function
       
    Dim numeros As String
    numeros = Replace(Replace(cep, "-", ""), ".", "")
    
    If Len(numeros) <> 8 Or _
    numeros = String(8, Left(numeros, 1)) Then
        MsgBox "CEP inválido", vbExclamation
        Exit Function
    End If
    
    EhCepValido = numeros
End Function

Private Function ObterCepXml(ByVal cepValido As String) As Object
    Set ObterCepXml = Nothing
    
    Dim url As String
    url = "https://viacep.com.br/ws/[cep]/xml/"
    url = Replace(url, "[cep]", cepValido)
    
    Dim inicio As Single
    inicio = 0
    
    Dim resposta As String
    resposta = ""
    
    On Error GoTo TratarErro
    
    With oReq
        .Open "GET", url, True
        .setRequestHeader "Accept", "application/xml"
        .send
        
        inicio = Timer
        
        While .readyState <> 4
            DoEvents
            
            If Timer - inicio > 30 Then
                Err.Raise vbObjectError + 1000, _
                    "Requisição API", _
                    "Tempo de espera pela resposta da API expirado"
            End If
        Wend
        
        If .Status <> 200 Then
            Err.Raise vbObjectError + 1001, _
                "Requisição API", _
                "Erro HTTP: " & .Status & vbNewLine & .statusText
            Exit Function
        End If
        
        Debug.Print .responseText
        
        resposta = .responseText
    End With
    
    If Not xmlDoc.LoadXML(resposta) Then
        Err.Raise vbObjectError + 1002, "Requisição API", "Erro ao carregar o XML"
    End If

    If xmlDoc.DocumentElement Is Nothing Then
        Err.Raise vbObjectError + 1003, "Requisição API", "A resposta não contém XML válido"
    End If
    
    On Error Resume Next
    If xmlDoc.SelectNodes("//erro").Length <> 0 Is Nothing Then
        Err.Raise vbObjectError + 1004, "Requisição API", "CEP não encontrado na base da API"
    End If
    On Error GoTo 0
    
    Set ObterCepXml = xmlDoc.SelectNodes("xmlcep")
    
    Exit Function

TratarErro:
    Set xmlDoc = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub PreencherCepAutomaticamente(ByVal form As MSForms.UserForm)
    Dim cep As String
    cep = form.Controls("txCep").Value
    
    Dim cepValido As String
    cepValido = EhCepValido(cep)
    
    If cepValido = "" Then Exit Sub

    On Error GoTo TratarErro
    
    Dim nodes As Object
    Set nodes = ObterCepXml(cep)
    
    On Error GoTo 0
    
    On Error Resume Next
    With form
        .Controls("txEndereco").Value = xmlDoc.SelectSingleNode("//logradouro").Text
        .Controls("txBairro").Value = xmlDoc.SelectSingleNode("//bairro").Text
        .Controls("txCidade").Value = xmlDoc.SelectSingleNode("//localidade").Text
        .Controls("txEstado").Value = xmlDoc.SelectSingleNode("//uf").Text
        .Controls("txRegiao").Value = xmlDoc.SelectSingleNode("//regiao").Text
        .Controls("txCodIbge").Value = xmlDoc.SelectSingleNode("//ibge").Text
        .Controls("txDDD").Value = xmlDoc.SelectSingleNode("//ddd").Text
    End With
    On Error GoTo 0

    Exit Sub

TratarErro:
    MsgBox Err.Source & vbCrLf & Err.Description, vbCritical
End Sub



