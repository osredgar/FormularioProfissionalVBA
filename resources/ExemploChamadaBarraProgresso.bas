Attribute VB_Name = "ExemploChamadaBarraProgresso"
Option Explicit
Option Private Module

Dim oBarra As clsBarraProgresso2026

Public Sub TesteBarraProgresso2026()
    
    Set oBarra = New clsBarraProgresso2026
    
    Dim x As Long
    x = 15000
    
    Dim i As Long
    For i = 1 To x
        oBarra.Atualizar i / x, "Teste"
    Next i
    
    Set oBarra = Nothing
    
End Sub
