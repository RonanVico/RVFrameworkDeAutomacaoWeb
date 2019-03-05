Attribute VB_Name = "EXEMPLOS_1"
Option Explicit




Public Sub VaAoGooglePesquiseWikipedia()
    'O nome da sub ja diz o que o exemplo faz ! rsrsrs
    Dim ie As New ieRV
    With ie
        Debug.Print "Quantos ies abertos ? "; .quantos_ies_abertos
        'inicia ie Invisivel
        Call .iniciaIE(SW_HIDE, True, InternetExplorer)
        Debug.Print "Quantos ies abertos ? "; .quantos_ies_abertos
        'navega , dighita wikipedia e pesquisa
        Call .NAVEGAR("www.google.com.br", SW_HIDE)
        Call .getElement(20, "tagname", "input", "title", "pesquisar").setAttribute("innerText", "Wikipedia")
        Call .getElement(20, "tagname", "input", "value", "*pesquisa*", "parentNode.tagname", "CENTER").Click
        .ie.visible = True
    End With
End Sub


 
 
Public Sub TabelaParaRangeDoExcel()
    Const HTMLtabelaExemplo As String _
        = "<table>" & _
              "<tr> " & _
                "<th>Month</th>" & _
                "<th>Savings</th>" & _
              "</tr>" & _
              "<tr>" & _
               " <td>January</td>" & _
              "  <td>$100</td>" & _
             " </tr>" & _
             "<tr>" & _
               " <td>Feb</td>" & _
              "  <td>$400</td>" & _
             " </tr>" & _
            "</table>"
    Dim ie As New ieRV
    With ie
        Call .iniciaIE
        Call .NAVEGAR("")
        .ie.Document.body.innerHTML = HTMLtabelaExemplo
        Call .TableToRange(.getElement(5, "tagname", "table"), ThisWorkbook.Sheets(1).Cells(6, 6))
    End With


End Sub


Public Sub BrincandoComPropsDoIE()
    Dim ie As New ieRV
    With ie
        .iniciaIE
        .NAVEGAR ("about:blank")
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=False, StatusBar:=False, Toolbar:=False, TheatherMode:=False, visible:=False)
        Call .wait(1000)
        Call .TrazerIeParaFrente
        Call .setPropertiesIES("", AddressBar:=True, MenuBar:=False, StatusBar:=False, Toolbar:=False, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .TrazerIeParaFrente
        Call .setPropertiesIES("", AddressBar:=True, MenuBar:=False, StatusBar:=False, Toolbar:=False, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .TrazerIeParaFrente
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=True, StatusBar:=False, Toolbar:=False, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .TrazerIeParaFrente
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=False, StatusBar:=True, Toolbar:=False, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .TrazerIeParaFrente
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=False, StatusBar:=False, Toolbar:=True, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .TrazerIeParaFrente
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=False, StatusBar:=False, Toolbar:=False, TheatherMode:=True, visible:=True)
        Call .wait(1000)
        Call .TrazerIeParaFrente
        Call .setPropertiesIES("", AddressBar:=False, MenuBar:=False, StatusBar:=False, Toolbar:=False, TheatherMode:=False, visible:=True)
        Call .wait(1000)
        Call .TrazerIeParaFrente
        Call .setPropertiesIES
    End With
End Sub


Public Sub AlterandoRegistrosImportantesDoIe()
    Dim ie As New ieRV
    '\/ Altera alguns registros do IE bacanas , dêem uma estudada
    Call ie.RegistryIE
End Sub


Public Sub Utilizando_IE()
    Dim ie As New ieRV
    With ie
        .iniciaIE
        .NAVEGAR ("https://www.linkedin.com/in/ronan-vico/")
        .iniciaIE SW_SHOWNORMAL, False
        .NAVEGAR ("https://github.com/RonanVico/")
    End With
End Sub




