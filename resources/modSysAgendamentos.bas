Attribute VB_Name = "modSysAgendamentos"
Option Explicit
Option Private Module

Public Function tbHorarios() As ListObject
    Set tbHorarios = shHorarios.ListObjects(1)
End Function

Public Function tbClientes() As ListObject
    Set tbClientes = shClientes.ListObjects(1)
End Function

Public Function tbServicos() As ListObject
    Set tbServicos = shServicos.ListObjects(1)
End Function

Public Sub AbrirHorarios()
    shHorarios.Activate
End Sub

Public Sub AbrirClientes()
    shClientes.Activate
End Sub

Public Sub AbrirServicos()
    shServicos.Activate
End Sub

Public Sub AbrirDashboard()
    shDashboard.Activate
End Sub
