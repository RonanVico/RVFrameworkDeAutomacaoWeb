Attribute VB_Name = "EXEMPLOS_USANDO_JAVASCRIPT"
Option Explicit


'Eu costumava a programar mais automações atraves do javascript / jquery das paginas do que
'utilizando os proprios elementos dentro do vba , pois era mais facil usar a janela de logs e dev
'dentro dos navegadores do que ir programando diretamente no vba,
'por exemplo , ao utilizar .ExecScript , era possivel usar comandos em Jquery ,
'se estou falando grego pra você , ignore provavelmente não é a hora de você saber disso!


'O mesmo exemplo no modulo exemplos_1 , só que diferente , em comentario está como foi feito
'sem o javascript , e como foi feito usando o javascript
Public Sub VaAoGooglePesquiseWikipedia_Apenas_Com_Javascripto()
    'O nome da sub ja diz o que o exemplo faz ! rsrsrs
    Dim ie As New ieRV
    With ie
        Call .iniciaIE
        Call .NAVEGAR("www.google.com.br")
        Call .waitElem("document.getElementsByName('q').item(0)", ".innerText = 'Wikipedia'", 20)
        'Call ie.getElement(20, "tagname", "input", "title", "pesquisar").setAttribute("innerText", "Wikipedia")
        Call .waitElem("document.getElementsByName('btnK').item(0)", ".click()", 20)
        'Call ie.getElement(20, "tagname", "input", "value", "*pesquisa*", "parentNode.tagname", "CENTER").Click
        Stop
    End With
End Sub


Public Sub AceitarUmAlerta()
    'LEIA OS COMENTARIOS ANTES DE DAR f5 Cabeçudo
    Dim ie As New ieRV
    With ie
        .iniciaIE noAddOns:=True
        Call .NAVEGAR("about:blank")
        'Essa linha lança um popup na janela do internet explorer
        'Call .execScript("setTimeout(""alert('ESSE É UM ERRO QUE O RONAN VICO CRIOU!');"", 1)")
        Call .ExecScriptAssync("alert('ESSE É UM ERRO QUE O RONAN VICO CRIOU!')", 1)
        Call .wait(1000)
        'Essa linha aceita o popup , lançando um erro com a mensagem que continha no popup.
        'Você pode também Mandar Falso e não receber o erro , apenas aceitando o alerta ex.: Call .aceitaAlerta(false)
        'Só sobe erro quando possuir o alerta , o erro é 12345 ,podendo ser tratado.
        Call .aceitaAlerta(True)
    End With
    Stop
End Sub
