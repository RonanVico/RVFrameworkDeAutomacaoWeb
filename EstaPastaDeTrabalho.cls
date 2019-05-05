VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
On Error GoTo fim
    Dim voice As Object
    Dim say As String
    Dim i As Long
    Set voice = CreateObject("SAPI.SpVoice")
    voice.Rate = 1
    voice.volume = 99
    
    say = "Olá , bem vindo ao Framework de automação web, criado por Ronan Vico" & _
          ", se possivel , acesse meu linkedin e ou meu github e me recomende! Obrigado! Aproveite e faça bom uso."
  
    voice.Speak say
    Call Utilizando_IE
fim:
End Sub


