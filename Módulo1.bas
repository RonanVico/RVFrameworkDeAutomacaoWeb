Attribute VB_Name = "Módulo1"
Option Explicit

'Dia 05/05/2019 , Realizei um video no youtube ensinando como usa-lo
'Link para o canal : https://www.youtube.com/user/RonanVico



Public Sub PegarTabelaDoGoogle()
    Dim ie As New ieRV
    
    With ie
        .iniciaIE SW_HIDE
        .NAVEGAR "www.google.com.br", SW_HIDE
        .getElement(5, "name", "q", "title", "Pesquisar").innerText = "Tabela do Brasileirão"
        .getElement(5, "name", "BtNk").Click
        Call .TableToRange(.getElement(5, "tagname", "table", "className", "liveresults-sports-immersive__stbl"), WSInicio.Range("A10"))
        .closeAllIE
    End With
End Sub




Public Sub BaixarUmaImagemAleatoria()
    Dim ie As New ieRV
    Dim palavaraMagica As String
    
    palavaraMagica = "Meme do Thanos em Português"
    With ie
        .iniciaIE SW_SHOWNORMAL
        .NAVEGAR "www.google.com.br", SW_SHOWNORMAL
        .getElement(5, "name", "q", "title", "Pesquisar").innerText = palavaraMagica
        .getElement(5, "name", "BtNk").Click
        .getElement(5, "className", "q qs", "innertext", "Imagens").Click
        .getElement(5, "tagname", "img", "alt", "Resultado de imagem*").Click
        Call URLDownloadToFile(0, .getElement(5, "id", "irc_mi").getAttribute("href"), VBA.Environ("TEMP") & "\Thanos.jpg", 0, 0)
        .closeAllIE
    End With
End Sub

