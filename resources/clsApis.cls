VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'DEVE-SE ADICIONAR O MODULO (modJson) QUE FAZ TODA A CONVERSAO DO JSON PARA O VBA
'ULTIMA ATUALIZACAO 21/02/2024 - EDGAR O.
'====================================================================================================

Private oJson As Object
Private oReq As Object
Private apiUrl As String
Private apiResp As String

Private Sub Class_Initialize()
    Set oReq = CreateObject("MSXML2.XMLHTTP")
End Sub

Private Sub Class_Terminate()
    Set oJson = Nothing
    Set oReq = Nothing
End Sub


'====================================================================================================
'FUNCIONALIDADE: CONSULTAR CEP ATRAVES DA API VIA CEP - "https://viaCep.com.br/ws/[CEP]/json/"
'Parametros: CEP, FORMULARIO
'RETORNO: -
'OBSERVACAO: MAXIMO 3 CONSULTAS POR MINUTO
'====================================================================================================
Public Sub ConsultarCEP( _
                            ByVal pCep As String, _
                            ByVal pForm As Object)
    
    On Error GoTo TratarErro

ObterSomenteNumeros:
    Let pCep = Replace(pCep, ".", "")
    Let pCep = Replace(pCep, "-", "")
    Let pCep = Replace(pCep, "_", "")
    
ValidarParametros:
    If (Len(pCep) = 0) Then GoTo Finalizar
    If (Len(pCep) < 8) Then GoTo TratarErroAPI
    Let apiUrl = "https://viaCep.com.br/ws/[CEP]/json/"
    Let apiUrl = Replace(apiUrl, "[CEP]", pCep)

ConsultarAPI:
    With oReq
        Let apiUrl = "https://viaCep.com.br/ws/[CEP]/json/"
        Let apiUrl = Replace(apiUrl, "[CEP]", pCep)
        Call .Open("GET", apiUrl, True)
        Call .setRequestHeader("Content-Type", "application/json")
        Call .send
ObterResposta:
        While (.readyState <> 4): DoEvents: Wend
        If (.Status <> 200) Then GoTo TratarErroConexao
        Let apiResp = .responseText
    End With
    
    
PreencherFormulario:
    On Error Resume Next
    With pForm
        Let .Controls("logradouro").Value = StrConv(oJson("logradouro"), 1)
        Let .Controls("bairro").Value = StrConv(oJson("bairro"), 1)
        Let .Controls("municipio").Value = StrConv(oJson("localidade"), 1)
        Let .Controls("uf").Value = StrConv(oJson("uf"), 1)
        Let .Controls("ibge").Value = StrConv(oJson("ibge"), 1)
        Let .Controls("ddd").Value = StrConv(oJson("ddd"), 1)
    End With
    On Error GoTo 0

Finalizar:
    On Error Resume Next
    Set oJson = Nothing
    Set pForm = Nothing
    Exit Sub

TratarErro:
    GoTo Finalizar

TratarErroConexao:
    Call MsgBox("TratarErro ao conectar a API." & vbNewLine & "Código do TratarErro: " & oReq.Status & vbNewLine & oReq.statusText, vbCritical + vbOKOnly, "Consultar Cep")
    GoTo Finalizar

TratarErroAPI:
    Call MsgBox("CEP Inválido", vbExclamation, "Consultar Cep")
    GoTo Finalizar
End Sub


'====================================================================================================
'FUNCIONALIDADE: CONSULTAR CNPJ ATRAVES DA API RECEITAWS - "https://receitaws.com.br/v1/Cnpj/[CNPJ]"
'Parametros: CNPJ, FORMULARIO
'RETORNO: -
'OBSERVACAO: MAXIMO 3 CONSULTAS POR MINUTO
'====================================================================================================
Public Sub ConsultarCnpj( _
                            ByVal pCnpj As String, _
                            ByVal pForm As Object)
    
    On Error GoTo TratarErro
    
ObterSomenteNumeros:
    Let pCnpj = Replace(pCnpj, ".", "")
    Let pCnpj = Replace(pCnpj, "-", "")
    Let pCnpj = Replace(pCnpj, "/", "")
    Let pCnpj = Replace(pCnpj, "_", "")
    
ValidarParametros:
    If (Len(pCnpj) = 0) Then GoTo Finalizar
    If (Len(pCnpj) < 14) Then GoTo TratarErroAPI
    
ConsultarAPI:
    With oReq
        Let apiUrl = "https://receitaws.com.br/v1/Cnpj/[CNPJ]"
        Let apiUrl = Replace(apiUrl, "[CNPJ]", pCnpj)
        Call .Open("GET", apiUrl, True)
        Call .setRequestHeader("Content-Type", "application/json")
        Call .send
ObterResposta:
        While (.readyState <> 4): DoEvents: Wend
        If (.Status <> 200) Then GoTo TratarErroConexao
        Let apiResp = .responseText
    End With
       
   
PreencherFormulario:
'   Verifica se é um cadastro ativo
    If (StrComp(oJson("situacao"), "ATIVA", 1) <> 0) Then
        Call MsgBox("CNPJ inativo, não pôde ser importado", vbExclamation, "Consultar Cnpj")
        GoTo Finalizar
    End If

'   Preenche os campos necessários para o formulário
    On Error Resume Next
    With pForm
        Select Case Len(Trim(oJson("fantasia")))
            Case 0:
                Let .Controls("fantasia") = StrConv(oJson("nome"), 1)
            Case Else:
                Let .Controls("fantasia") = StrConv(oJson("fantasia"), 1)
        End Select
        Let .Controls("nome").Value = StrConv(oJson("nome"), 1)
        Let .Controls("cep").Value = StrConv(oJson("cep"), 1)
        Let .Controls("numero").Value = Val(oJson("numero"))
        Let .Controls("logradouro").Value = StrConv(oJson("logradouro"), 1)
        Let .Controls("bairro").Value = StrConv(oJson("bairro"), 1)
        Let .Controls("municipio").Value = StrConv(oJson("municipio"), 1)
        Let .Controls("uf").Value = StrConv(oJson("uf"), 1)
        Let .Controls("telefone").Value = StrConv(oJson("telefone"), 1)
        Let .Controls("abertura").Value = StrConv(oJson("abertura"), 1)
        Let .Controls("porte").Value = StrConv(oJson("porte"), 1)
    End With
    On Error GoTo 0

Finalizar:
    On Error Resume Next
    Set oJson = Nothing
    Set pForm = Nothing
    Exit Sub

TratarErro:
    GoTo Finalizar

TratarErroConexao:
    Select Case oReq.Status
        Case 429, 504
            Call MsgBox("Esta API permite 3 consultas por minuto, tente novamente mais tarde" & vbNewLine & "Código do Erro: " & oReq.Status & vbCr & oReq.statusText, vbCritical + vbOKOnly, "Consultar Cnpj")
        Case Else
            Call MsgBox("Erro ao conectar a API" & vbNewLine & "Código do Erro: " & oReq.Status & vbCr & oReq.statusText, vbCritical + vbOKOnly, "Consultar Cnpj")
    End Select
    GoTo Finalizar

TratarErroAPI:
    Call MsgBox("CNPJ Inválido", vbExclamation, "Consultar Cnpj")
    
End Sub

'Public Sub BuscarFeriadosXml(ByVal pCodIbge As String, ByVal pAno As Integer)
'
'DefinirVariaveis:
'    Dim mUrl As String, mSql As String
'    Dim req As MSXML2.XMLHTTP60
'    Set req = New MSXML2.XMLHTTP60
'    Dim xDoc As MSXML2.DOMDocument60
'    Set xDoc = New MSXML2.DOMDocument60
'    Dim xNodes As MSXML2.IXMLDOMNodeList
'    Dim xNode As MSXML2.IXMLDOMNode
''    Set oForm = FR_Menu
''    Let oForm.Caption = oForm.Caption & "... Atualizando lista de feriados"
'
'DefinirUrl:
'    Let mUrl = "https://api.calendario.com.br/?ano=[ANO]&ibge=[IBGE]&token=[TOKEN]"
'    Let mUrl = Replace(mUrl, "[ANO]", pAno)
'    Let mUrl = Replace(mUrl, "[IBGE]", pCodIbge)
'    Let mUrl = Replace(mUrl, "[TOKEN]", "ZWRnYXJvdGF2aW8uc3JAZ21haWwuY29tJmhhc2g9MTE4NTQ3Njc3")
'
'ExecutarChamadaApi:
'    On Error GoTo TratarErro
'
'    Call req.Open("GET", mUrl, False)
'    Call req.SetRequestHeader("Content-Type", "application/xml")
'    Call req.Send
'
'    While (req.readyState <> 4)
'       DoEvents
'    Wend
'
'    If (req.Status <> 200) Then
'        Call MsgBox(req.statusText, vbCritical + vbOKOnly, req.Status)
'        GoTo Finalizar
'    End If
'
''   Limpar tabela de feriados
'    Set oBD = New clsBancoDados
'    With oBD
'        Let mSql = "DELETE FROM FERIADOS;"
'        Call .ConectarBancoDados
'        Call .ExecutarQuery(mSql)
'        Call .DesconectarBancoDados
'    End With
'
'    On Error GoTo 0
'ObterRespostaApi:
'    On Error GoTo TratarErro
'
'    Call xDoc.LoadXML(req.ResponseText)
'
'    Set xNodes = xDoc.SelectNodes("/events/event")
'
'InserirValoresBancoDados:
'    For Each xNode In xNodes
'        With xNode
'            If (.SelectSingleNode("type_code").Text < 4) Then
'                Let mSql = ""
'                Let mSql = mSql & "INSERT INTO FERIADOS(DTFERIADO,FERIADO,TIPO)"
'                Let mSql = mSql & " VALUES('@DTFERIADO@','@FERIADO@','@TIPO@');"
'                Let mSql = Replace(mSql, "@DTFERIADO@", .SelectSingleNode("date").Text)
'                Let mSql = Replace(mSql, "@FERIADO@", .SelectSingleNode("name").Text)
'                Let mSql = Replace(mSql, "@TIPO@", .SelectSingleNode("type").Text)
'                With oBD
'                    Call .ConectarBancoDados
'                    Call .ExecutarQuery(mSql)
'                    Call .DesconectarBancoDados
'                End With
'            End If
'        End With
'    Next xNode
'
'    Call MsgBox("Lista de feriados atualizada", vbInformation)
'
'    On Error GoTo 0
'
'Finalizar:
'    On Error Resume Next
''    Let oForm.Caption = Replace(oForm.Caption, "... Atualizando lista de feriados", "")
''    Set oForm = Nothing
'    Set req = Nothing
'    Set xNodes = Nothing
'    Set xNode = Nothing
'    Set xDoc = Nothing
'    Set oBD = Nothing
'    Exit Sub
'
'TratarErro:
'    GoTo Finalizar
'
'End Sub




