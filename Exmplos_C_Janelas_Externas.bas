Attribute VB_Name = "Exmplos_C_Janelas_Externas"
Option Explicit

Public Sub EnviarArquivoParaAlgumSite()
  Dim fileFullName      As String
  Const html As String = " <html>" _
        & "<!-- saved from url=(0017)http://localhost/ --> " _
        & "<form method=||post|| action=||nothing|| enctype=||multipart/form-data||>" _
        & "<input type=||hidden|| name=||method|| value=||post||>" _
        & "Your Key:" _
        & "<input type=||text|| name=||key|| value=||none||>" _
        & "The file:" _
        & "<input type=||file|| name=||file|| value = ||||>" _
        & "<input type=||submit|| value=||Upload and get the ID||>" _
         & "</form> "

    Dim ie As New ieRV
    With ie
        .iniciaIE
        .ie.visible = True
        .ie.Document.body.innerHTML = VBA.Replace(html, "||", """")
        
        fileFullName = VBA.Environ("temp") & "\ts.txt"
        Open fileFullName For Output As #1
            Write #1, "T"
        Close #1
        
        'Tem que ser assyncrono para o IE não travar ! Viu como um framework é bacana
        Call .ExecScriptAssync("document.getElementsByName('file')(0).click()", 5)
        Call .SendToWindowOpen(VBA.Environ("temp") & "\ts.txt")
    End With
    
End Sub

Public Sub SalvarAlgoQuePossuiJanelaDeSalvar()
    Dim ie As New ieRV
    With ie
        .iniciaIE
        .NAVEGAR ("https://www.linkedin.com/in/ronan-vico/")
        .TrazerIeParaFrente
        .TrazerIeParaFrente
        .wait (5000)
        .TrazerIeParaFrente
        .TrazerIeParaFrente
        VBA.SendKeys ("^s")
        VBA.SendKeys ("^s")
        Call .SendToWindowSaveAs(ThisWorkbook.Path & "\exemplo", "Salvar página da Web")
    End With
End Sub


Public Sub ImprimirAlgoComJanelaDeImprimir()
    Dim ie As New ieRV
    With ie
        .SetPrintPDF ("Microsoft Print to PDF") '<- Altere conforme seu gosto
        .iniciaIE
        .NAVEGAR ("https://www.linkedin.com/in/ronan-vico/")
        .execScript ("window.print()")
        .wait (1000)
        .SendEnterToSaveOrOpenWindow
        Call .SendToWindowSaveAs(ThisWorkbook.Path & "\exemplo", "Salvar Saída de Impressão como")
    End With
End Sub



