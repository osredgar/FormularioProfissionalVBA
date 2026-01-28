Attribute VB_Name = "modGlobal"
Option Explicit
Option Private Module


'Dados do aplicativo
Public Const APP_NOME As String = "Transfermarkt Scraping PowerQuery"


'Formatos de numeros
Public Const FORMATO_PORCENTAGEM As String = "#.##0,0%;[Cor10]-#.##0,0%;#.##0,0%;@"
Public Const FORMATO_DECIMAL As String = "#.##0,00;[Cor10]-#.##0,00;#.##0,00;@"
Public Const FORMATO_INTEIRO As String = "#.##0;[Cor10]-#.##0;#.##0;@"
Public Const FORMATO_DATA As String = "mm/aaaa"
Public Const FORMATO_DATA_EXTENSO As String = "mmm/aa"


'Fontes
Public Const FONTE_TAMANHO As Long = 9
Public Const FONTE_CABECALHO As String = "Segoe UI Semibold"
Public Const FONTE_TEXTO As String = "Segoe UI"


'Cores BootStrap
Public Const COR_NONE As Long = -4142 'xlNone
Public Const COR_BRANCO As Long = 16447992 'light
Public Const COR_FONTES As Long = 2696481 'dark
Public Const COR_SUCESSO As Long = 3119151 'sucess
Public Const COR_ATENCAO As Long = 184319 'warning
Public Const COR_PERIGO As Long = 1458169 'danger

Public Const COR_AZUL As Long = 16608781 'blue500
Public Const COR_CINZA As Long = 12432813 'gray500
Public Const COR_VERDE As Long = 5539609 'green500
Public Const COR_VERMELHO As Long = 4535772 'red500
Public Const COR_AMARELO As Long = 508415 'yellow500
Public Const COR_CIANO As Long = 15780365 'cyan500
Public Const COR_INDIGO As Long = 15863910 'indigo500
Public Const COR_LARANJA As Long = 1343229 'orange500

Public Const ALERTA_AZUL As Long = 16689262 'blue300
Public Const ALERTA_CINZA As Long = 15131358 'gray300
Public Const ALERTA_VERDE As Long = 10008437 'green300
Public Const ALERTA_VERMELHO As Long = 9406186 'red300
Public Const ALERTA_AMARELO As Long = 7002879 'yellow300
Public Const ALERTA_CIANO As Long = 16179054 'cyan300
Public Const ALERTA_INDIGO As Long = 16216227 'indigo300
Public Const ALERTA_LARANJA As Long = 7516926 'orange300




