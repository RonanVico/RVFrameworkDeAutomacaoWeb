VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ieRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'BIBLIOTECAS NECESSARIAS
'Microsoft Internet Controls
'Microsoft HTML Object Library



'Peço porfavor , para me ajudarem com uma recomendação na pagina do linkedin e ou  um Star no github , ajuda bastante pessoal!
'https://www.linkedin.com/in/ronan-vico/
'https://github.com/RonanVico/

' ----------------------------------------------------------------
' Purpose: Framework criado para ajudar na automação web
' levei cerca de 1 ano para aprender tudo que foi colocado no framework , muita coisa foi retirada
' por inutlidade ou por melhorias que subistituiram códigos antigos.
' Author: Ronan Vico
' ----------------------------------------------------------------


Public ie As InternetExplorer
Private IE_URL    As String
Private LOG     As String

'TIPOS DE AMOSTRAGEM DA JANELA DO INTERNET EXPLORER
Public Enum IE_WINDOW_SHOW
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
End Enum

Public Enum NAVIGATION_OPTIONS
    navOpenInNewWindow = 1
    navNoHistory = 2
    navNoReadFromCache = 4
    navNoWriteToCache = 8
    navAllowAutosearch = 16
    navBrowserBar = 32
    navHyperlink = 64
    navEnforceRestricted = 128
    navNewWindowsManaged = 256
    navUntrustedForDownload = 512
    navTrustedForActiveX = 1024
    navOpenInNewTab = 2048
    navOpenInBackgroundTab = 4096
    navKeepWordWheelText = 8192
    navVirtualTab = 16384
    navBlockRedirectsXDomain = 32768
    navOpenNewForegroundTab = 65536
End Enum


Public Enum TARGETFRAME_OPTIONS
      blank = 1
      Parent = 2
      self = 3
      top = 4
End Enum

Public Enum IE_TYPE
    InternetExplorer = 1
    InternetExplorerMedium = 2
End Enum

' ----------------------------------------------------------------
' Procedure Name: iniciaIE
' Purpose: garante a inicialização do internet explorer utilizando alguns parametros
' Parameter SHOW_IE (IE_WINDOW_SHOW): Como mostar o Ie
' Parameter KILL_ALL_OTHERS_IE (Boolean): Matar todos os IES ja abertos
' Parameter IE_TYPE (IE_TYPE): Pode ser IE ou IEMedium
' Parameter noAddOns (Boolean): Inicializa o IE sem complementos
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Sub iniciaIE(Optional ByVal SHOW_IE As IE_WINDOW_SHOW = IE_WINDOW_SHOW.SW_SHOWNORMAL, Optional ByVal KILL_ALL_OTHERS_IE As Boolean = True, Optional ByVal IE_TYPE As IE_TYPE, Optional ByVal noAddOns As Boolean = False)

    'Mata os explorer vivo , sim deixar 2 vezes.
    If KILL_ALL_OTHERS_IE Then
        On Error Resume Next
            ie.Quit
            Set ie = Nothing
        On Error GoTo 0
        Call closeAllIE
    End If
    
    Dim tentativas As Integer
    tentativas = 0
    
    On Error GoTo TESTE_ERR
      
    Select Case IE_TYPE
      Case 2
        Set ie = New InternetExplorerMedium
      Case Else
        Set ie = New InternetExplorer ' CreateObject("InternetExplorer.Application")
        'Set ie = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")
    End Select
    
    
    WriteLog ("ABRINDO IE...")

    ShowWindow ie.hwnd, SHOW_IE
    
    'ie.visible = visible
    ie.Silent = True
    If noAddOns Then
        ie.Navigate "about:NoAdd-ons"
    Else
        ie.Navigate "about:blank"
    End If

    Call aceitaIE11Alert
    WriteLog ("NAVEGANDO PARA O SITE...")

    Exit Sub
    
TESTE_ERR:
    If Err.Number = 429 Or InStr(Err.Description, "Erro de automação") <> 0 Then
        If tentativas > 25 Then Exit Sub
        tentativas = tentativas + 1
        WriteLog ("FECHANDO IES ABERTOS...")
        closeAllIE
        'Call taskkillSafe("iexplore.exe", True)
        Call killProcess("iexplore.exe")
        WriteLog ("RETORNANDO A iniciação do Internet explorer!")
        wait (2000)
        Resume
    ElseIf tentativas < 2 Then
        'Call taskkillSafe("iexplore.exe", True)
        closeAllIE
        wait (30000)
        tentativas = tentativas + 1
        Resume
    End If
    
'    Stop
'        MsgBox err.Description & err.Number
'    Resume
End Sub

' ----------------------------------------------------------------
' Procedure Name: NAVEGAR
' Purpose: Navegar o IE , Só que de um jeito mais facil de entender
' Parameter url (String): URL que você quer ir
' Parameter SHOW_IE (IE_WINDOW_SHOW): como você qer mostrar o IE Hide ou maximized ou etc
' Parameter navOptionsPodeSomarOsTipos (NAVIGATION_OPTIONS): Opções de navegações podendo ser somadas
' Parameter targetFrame (TARGETFRAME_OPTIONS): Target
' Parameter postData (Variant): Isso é um pouco mais avançado mas é o postdata do request
' Parameter Headers (Variant): header do request
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Sub NAVEGAR(ByVal url As String, Optional ByVal SHOW_IE As IE_WINDOW_SHOW = IE_WINDOW_SHOW.SW_SHOWNORMAL, Optional ByVal navOptionsPodeSomarOsTipos As NAVIGATION_OPTIONS = 1, Optional ByVal targetFrame As TARGETFRAME_OPTIONS, Optional ByVal postData As Variant, Optional ByVal Headers As Variant)

  On Error GoTo RIP
  Dim stargetFrame As String
  Select Case targetFrame
      Case TARGETFRAME_OPTIONS.blank
          stargetFrame = "_blank"
      Case TARGETFRAME_OPTIONS.Parent
          stargetFrame = "_parent"
      Case TARGETFRAME_OPTIONS.top
          stargetFrame = "_top"
      Case Else
        stargetFrame = "_self"
  End Select
  
    ie.Navigate2 url, navOptionsPodeSomarOsTipos, stargetFrame, postData, Headers
    
    Do While ie.ReadyState = 4: DoEvents: Loop   'Do While
    ShowWindow ie.hwnd, SHOW_IE
    Call aceitaIE11Alert
    Do Until ie.ReadyState >= 3: DoEvents: Call aceitaAlerta(False): Loop  'Do Until
    Exit Sub
    Resume
RIP:

    Select Case Err.Number
      Case 0
      Case 91
          MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NAVEGAR, " & vbNewLine & _
            "INICIE O INTERNET EXPLORER! OBJETO NÃO ENCONTRADO " & Erl & ".", vbCritical
      Case Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NAVEGAR, line " & Erl & "."
    End Select
End Sub





' ----------------------------------------------------------------
' Procedure Name: closeAllIE
' Purpose: Função que feixa todos os IEs aberto ou então só o da URL.
' Parameter url (String): url se vc quiser fechar só o q CONTEM a url
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Function closeAllIE(Optional ByVal url As String = "")
On Error GoTo fim
    'Call CreateObject("WScript.Shell").Run("taskkill /F /IM iexplore.exe /t /FI ""USERNAME eq %USERNAME% ""  ", 0, True)
    Dim MyWindows As New SHDocVw.ShellWindows
    Dim ie As SHDocVw.InternetExplorer
    
    For Each ie In MyWindows
        'Debug.Print Window.Name
        If ie.Name = "Internet Explorer" Then
            If ie.LocationURL <> "" Then
                If url = "" Then
                    ie.Quit
                    Set ie = Nothing
                ElseIf url = ie.LocationURL Then
                    ie.Quit
                    Set ie = Nothing
                End If
            End If
        End If
    Next ie
fim:
End Function

'ACRECENTA AO LOG EXISTENTE A NOVA MENSAGEM DE LOG
Public Sub WriteLog(ByVal LOGNEW As String)
  
    LOGNEW = Replace(Replace(LOGNEW, "'", ""), Chr(34), "")
    LOG = UCase(Replace(LOG, "...", "") + LOGNEW) + vbNewLine + vbNewLine
End Sub


' ----------------------------------------------------------------
' Procedure Name: aceitaIE11Alert
' Purpose:  'DEVIDO A UMA JANELA DE "INTERNET EXPLORER 11" SOBRE SEGURANÇA APARECER
'            'È NECESSARIO FECHAR ESSA PORCARIA , JA Q EU NAO ACHEI O REGISTRO DESSE LIXO
' Author: Ronan Vico
' ----------------------------------------------------------------
Private Sub aceitaIE11Alert()
     Dim i As Long
     Dim hwnd&
     For i = 0 To 1400
        hwnd = FindWindowExA(0, 0, "#32770", "Internet Explorer 11")
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", vbNullString)
        'manda um ESC para fechar a janela maldita e segue o jogo
        PostMessageA hwnd, &H100, 27, 0          '
        PostMessageA hwnd, &H101, 27, 0          '
        PostMessageA hwnd, &H100, 27, 0          '
        PostMessageA hwnd, &H101, 27, 0          '
        DoEvents
    Next i
End Sub


' ----------------------------------------------------------------
' Procedure Name: aceitaAlerta
' Purpose:Aceita qualquer tipo de alerta ou popup , e da a opção de retornar como erro
' a mensagem contida nesse alerta!!
' Parameter RaiseError (Boolean): Sobe um erro onde a descrição é o texto da mensagem
' Return Type: Boolean
' Author: Ronan Vico
' Date: 17 10 2018 by RV
' ----------------------------------------------------------------
Public Function aceitaAlerta(Optional RaiseError As Boolean = True) As Boolean
    aceitaAlerta = False
    Dim ie_hwnd, frm_hwnd, n&, folder_bak$, file_name$, tmp_dir$
    Dim textoAlerta As String
    Dim hwndDentro As Long
    Dim textodedentro As String
    ' wait for the download dialogue (IEFrame/Frame Notification Bar/DirectUIHWND)
    frm_hwnd = FindWindowExA(0, 0, "#32770", "Mensagem da página da web")
    If frm_hwnd = 0 Then frm_hwnd = FindWindowExA(0, 0, "#32770", "Internet Explorer")
    
    If frm_hwnd <> 0 Then
        If WindowTextString(frm_hwnd) <> "" Then
            textoAlerta = WindowTextString(frm_hwnd) & vbNewLine
            hwndDentro = FindWindowExA(frm_hwnd, 0, "Static", vbNullString)
            textodedentro = WindowTextString(hwndDentro)
            textoAlerta = textoAlerta & vbNewLine & textodedentro
            textodedentro = WindowTextString(GetWindow(hwndDentro, 2))
            textoAlerta = textoAlerta & vbNewLine & textodedentro
            textodedentro = WindowTextString(GetWindow(hwndDentro, 5))
            textoAlerta = textoAlerta & vbNewLine & textodedentro
            
            PostMessageA frm_hwnd, &H100, 27, 0  '
            PostMessageA frm_hwnd, &H101, 27, 0  '
            PostMessageA frm_hwnd, &H100, 27, 0  '
            PostMessageA frm_hwnd, &H101, 27, 0  '
            aceitaAlerta = True
            If RaiseError = True Then
              Err.Raise 12345, , textoAlerta
            End If
        End If
    End If
End Function



' ----------------------------------------------------------------
' Procedure Name: killProcess
' Purpose: Mata um processo.
' Parameter processname (String): Por exemplo killProcess("excel.exe") ou killProcess("iexplore.exe")
' Return Type: Boolean
' Author: Ronan Vico
' Date: 05/03/2019
' ----------------------------------------------------------------
Public Function killProcess(ByVal processname As String) As Boolean
killProcess = False
Dim oServ As Object
Dim cProc As Variant
Dim oProc As Object
Dim UserName As String
Dim userDomain As String
On Error GoTo fim
    Set oServ = GetObject("winmgmts:")
    Set cProc = oServ.ExecQuery("Select * from Win32_Process where Name = " & Chr(34) & processname & Chr(34))
    
    On Error Resume Next
    For Each oProc In cProc
        Call oProc.getOwner(UserName, userDomain)
        If LCase(UserName) = LCase(VBA.Environ$("USERNAME")) Then
            oProc.Terminate
        End If
    Next
    If Err.Number = 0 Then killProcess = True
fim:
    
End Function


' ----------------------------------------------------------------
' Procedure Name: wait
' Purpose: Espera alguns segundos ,muito util na automação web.
' Parameter ms (Single):
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Sub wait(ms As Single)
    Dim SngSec As Single
    SngSec = Timer + (ms / 1000)
    
    Do While Timer < SngSec
        DoEvents
    Loop
End Sub



' ----------------------------------------------------------------
' Procedure Name: getIeWindowWhere
' Purpose: Procura e retorna uma Janela internet explorer que contenha no URL o parametro passado
' por exemplo se você abrir manualmente um Google.com.br , pode setar num objeto atraves dessa função.
' Quando usar essa função ? Quando janelas popups abrirem em outros Internet explorer e você consegue
' mudar o objeto para a outra janela , apenas utilizando o URl que ela contem.
' Parameter instrLocationURL (String):
' Return Type: InternetExplorer
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Function getIeWindowWhere(Optional ByVal instrLocationURL As String = "*") As InternetExplorer
Dim MyWindows As Object
Dim window As Variant
    Set MyWindows = CreateObject("SHELL.APPLICATION").Windows
    For Each window In MyWindows
        'Debug.Print Window.Name
        If window.Name = "Internet Explorer" Then
            If LCase(window.LocationURL) Like "*" & LCase(instrLocationURL) & "*" Then
                Set getIeWindowWhere = window
                ShowWindow window.hwnd, SW_SHOWMINIMIZED
                Exit Function
            End If
        End If
    Next window
    
    Set getIeWindowWhere = Nothing
End Function

' ----------------------------------------------------------------
' Procedure Name: disableAlerts
' Purpose: Desabilita a função alert() , util quando você nao qer cair num travamento
'  nem sempre funciona tem paginas q acabam não funcionando
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Sub disableAlerts()
    Call execScript("Window.prototype.alert = function(){}", , False)
End Sub

' ----------------------------------------------------------------
' Procedure Name: execScript
' Purpose:Executa um javascript no internet explorer , facil né
' Parameter script (String): O script em javascripto
' Parameter Lenguage (String): cara na vdd vc só vai uasr Javascript ignora isso
' Parameter RaiseError (Boolean): Sobe um err.raise se e for true
' Return Type: String
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Function execScript(ByVal script As String, Optional Lenguage As String = "JavaScript", Optional ByVal RaiseError As Boolean = True) As String
       If RaiseError = False Then On Error Resume Next
       execScript = ie.Document.parentWindow.execScript(script, Lenguage)
       If RaiseError = False Then On Error GoTo 0
End Function


' ----------------------------------------------------------------
' Procedure Name: isVisible
' Purpose:'Cria uma função em JS que retorna se o elemento esta visivel ou não
'            'Cria um elemento html pra jogar a resposta , joga a resposta da função
'            'visible do JS no elemento para depois pegar e retorna
' Parameter scriptGetElement (String):
' Return Type: Boolean
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Function isVisible(ByVal scriptGetElement As String) As Boolean
'Criada por Ronan Vico


On Error GoTo trataerro
    Dim scriptJS As String
    
    scriptJS = "function isVisibleVBAFUNCTION(elem) {" _
& "    if (!(elem instanceof Element)) throw Error('DomUtil: elem is not an element.'); " _
& "    const style = getComputedStyle(elem); " _
& "    if (style.display === 'none') return false; " _
& "    if (style.visibility !== 'visible') return false; " _
& "    if (style.opacity < 0.1) return false; " _
& "    if (elem.offsetWidth + elem.offsetHeight + elem.getBoundingClientRect().height + " _
& "        elem.getBoundingClientRect().width === 0) { " _
& "        return false; " _
& "    } " _
& "    const elemCenter   = { " _
& "        x: elem.getBoundingClientRect().left + elem.offsetWidth / 2, " _
& "        y: elem.getBoundingClientRect().top + elem.offsetHeight / 2 " _
& "    }; " _
& "    if (elemCenter.x < 0) return false; " _
& "    if (elemCenter.x > (document.documentElement.clientWidth || window.innerWidth)) return false; " _
& "    if (elemCenter.y < 0) return false; " _
& "    if (elemCenter.y > (document.documentElement.clientHeight || window.innerHeight)) return false; " _
& "    let pointContainer = document.elementFromPoint(elemCenter.x, elemCenter.y); " _
& "    do { " _
& "        if (pointContainer === elem) return true; " _
& "    } while (pointContainer = pointContainer.parentNode); " _
& "    return false; " _
& "} "
    
    Call execScript(ie, scriptJS)
    
    Dim uniqID As String
    'Cria elemento unico para pegarmos a resposta
    uniqID = "RONAN_VICO_ELEMENTO_VBA"
    If IsNull(ie.Document.getElementById(uniqID)) Then
        scriptJS = "(function(){var x = document.createElement('p'); x.id='" & uniqID & "';x.innerText=isVisibleVBAFUNCTION(" & scriptGetElement & ");document.body.appendChild(x); })()"
    Else
        scriptJS = "document.getElementById('" & uniqID & "').innerText=isVisibleVBAFUNCTION(" & scriptGetElement & ");"
    End If
    
    On Error Resume Next
    Call execScript(ie, scriptJS)
    'Pega a respsota da função
    isVisible = IIf(ie.Document.getElementById(uniqID).InnerText = "true", True, False)
    Exit Function
trataerro:
    Err.Raise Err.Number, , Err.Description
End Function



Public Sub PrintaHtmlNoWORD(ie As Object)
    'apenas para fins de programação , pegar o texto do documento durante o desenvolvimento
    Call CreateObject("WScript.Shell").Run("taskkill /F /IM winword.exe /t", 0, True)
    Dim doc As Object
    Dim wd As Object
    Set wd = CreateObject("Word.Application")
    wd.visible = True
    wd.Documents.Add
    Set doc = ie.Document
    wd.Documents(1).Content = doc.head.innerHTML
    wd.Documents(1).Content = wd.Documents(1).Content & doc.body.innerHTML
End Sub

Public Function criaElemento(ByVal tag As String, ByVal id As String) As String
       Call ie.Document.parentWindow.execScript("var elem =  document.body.appendChild(document.createElement('" & tag & "')); elem.id = '" & id & "';", "JavaScript")
       criaElemento = id
End Function

' ----------------------------------------------------------------
' Procedure Name: waitElem
' Purpose: Atraves de um Javascript esperamos o elemento aparecer na pagina e executamos alguma função ,
'                    por exemplo "document.getElementsByName(""name"")(0)" , <- Exemplo de javascript
' Parameter scriptGetElement (String): /\
' Parameter Action (String): ".value = ''" , ".click()" , ".innerText = '' "
' Parameter seconds (Integer): Segundos pra ficar esperando até aperecer
' Parameter RaiseError (Boolean): Subir erro ? Se não encontrar e for raiseError = true então vai dar Erro
' Parameter WaitPageComplete (Boolean): Esperar pagina ficar readystate complete
' Parameter waitVisibleTrue (Boolean): Esperar ficar visivel o elemento
' Return Type: Boolean
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Function waitElem(ByVal scriptGetElement As String, Optional ByVal Action As String = "", Optional ByVal seconds As Integer = 20, Optional ByVal RaiseError As Boolean = True, Optional ByVal WaitPageComplete As Boolean = True, Optional ByVal waitVisibleTrue As Boolean = False) As Boolean
waitElem = False
    Dim SecondsToWait As Date
    If seconds >= 60 Then seconds = 59
    If seconds <= 0 Then seconds = 1
    SecondsToWait = CDate("00:00:" & seconds)
    
    Dim agora As Date
    agora = DateTime.Time
    
    Dim i As Long
    On Error GoTo Err_handler
      Call ie.Document.parentWindow.execScript(scriptGetElement & ".tagName")
      waitElem = True
      If WaitPageComplete Then
          Do Until ie.ReadyState >= 3: DoEvents: Loop  'Do Until
      End If
      If waitVisibleTrue Then
         While isVisible(scriptGetElement) = False
            DoEvents
            Err.Raise 123, "WaitElem Function", "ELEMENTO PRESENTE MAS  NÃO VISIVEL!"
            Call wait(3000)
         Wend
      End If
      If Action <> "" Then
          If left(Action, 1) <> "." Then Action = "." & Action
          Call ie.Document.parentWindow.execScript(scriptGetElement & Action)
      End If
      Exit Function
      
Err_handler:
    If CDate(Format(Now - agora, "hh:mm:ss")) < SecondsToWait Then
      DoEvents
      If Err.Number = 123 Then
        Resume Next
      ElseIf Err.Number = 462 Then
        Err.Raise Err.Number, , Err.Description
      End If
      Resume
    End If
    If RaiseError Then Err.Raise "7", , "ELEMENTO NÃO ENCONTRADO " & scriptGetElement
End Function



' ----------------------------------------------------------------
' Procedure Name: RegistryIE
' Purpose: Altera alguns registros di internet explorer , é mais facil ler a função toda , pois
'       cada registro significa uma coisa
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Sub RegistryIE()
'Made BY RV
'Ajusta os registros do ie
   Dim a As Variant
    Dim shl As Object
    Set shl = Nothing
    
    Const IE_KEYS = "HKCU\Software\Microsoft\Internet Explorer"
    Const MAIN_KEYS = IE_KEYS & "\Main"
    
    If shl Is Nothing Then
        Set shl = CreateObject("WScript.Shell")
        'COLOCA O ZOOM em 100%
        shl.regWrite IE_KEYS & "\Zoom\ZoomFactor", 100000, "REG_DWORD"
        'shl.regWrite "HKCU\Software\Microsoft\Internet Explorer\Zoom\ResetZoomOnStartup", 1, "REG_DWORD" TESTAR
        'Não Salva no historico
        shl.regWrite MAIN_KEYS & "\Save_Session_History_On_Exit", "no", "REG_SZ"
        'Limpa o historico
        shl.regWrite MAIN_KEYS & "\ClearBrowsingHistoryOnExit", 1, "REG_DWORD"
        'Libera popup de janelas
        'shl.regWrite IE_KEYS & "\New Windows\PopupMgr", "0", "REG_DWORD"
        'TESTAR DESABILITA EXTENSõES
        shl.regWrite MAIN_KEYS & "\Enable Browser Extensions", "no", "REG_SZ"
        'Não avisar se o download acabou evitando janelas
        shl.regWrite MAIN_KEYS & "\NotifyDownloadComplete", "no", "REG_SZ"
        'tira a Urltoolbar de cima
        'shl.regWrite MAIN_KEYS & "\Show_URLToolBar", "no", "REG_SZ"
        'retira a toolbar
        'shl.regWrite MAIN_KEYS & "\Show_ToolBar", "no", "REG_SZ"
        'Mostra URL no status bar
        'shl.regWrite MAIN_KEYS & "\Show_URLinStatusBar", "no", "REG_SZ"
        'mostra o status bar
        'shl.regWrite MAIN_KEYS & "\Show_StatusBar", "no", "REG_SZ"
        'mostra url completa
        'shl.regWrite MAIN_KEYS & "\Show_FullURL", "no", "REG_SZ"
        'desabilita pergunta do "Deseja auto preencher"
        'shl.regWrite IE_KEYS & "\IntelliForms\AskUser", 0, "REG_DWORD"
        'Habilita alguns comandos no ie
        On Error Resume Next
        'Debug.Print "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BFCACHE\iexplorer.exe"
        'Deletando esse registro nos ajuda a deixar o IE mais rapidos. COnfia em MIM CARA
        shl.RegDelete MAIN_KEYS & "\TabProcGrowth"
        On Error GoTo 0
        
        'DESABILITA PROTECTED MODE
        'For i = 0 To 4
        '    shl.regWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" & i & "\2500", 3, "REG_DWORD"
        'Next i
        
        'For i = 2 To 4
        '    'Display mixed content = ENabled 0 , disable = 3
        '    shl.regWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" & i & "\1609", 0, "REG_DWORD"
        '    'shl.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" & i & "\1809", 0, "REG_DWORD"
            'shl.regwrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\" & i & "\1809", 0, "REG_DWORD"
        'Next i
    End If
End Sub


' ----------------------------------------------------------------
' Procedure Name: retiraAlertJS
' Purpose: Meio bosta esssa função mas ja ajudou em algumas vezes
' desligamos todos os alerts escrevendo alert errado assim af unçãio não ira funcionar
' Parameter htmlDoc (Object): documento html da onde queremos tirar o alertas
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Sub retiraAlertJS(htmlDoc As Object)
    On Error GoTo fim
    
    Dim i
    Dim frame
    If htmlDoc.getElementsByTagName("script").Length <> 0 Then
        For i = 0 To htmlDoc.getElementsByTagName("script").Length - 1
            If InStr(htmlDoc.getElementsByTagName("script")(i).innerHTML, "alert") Then
                '  Debug.Print htmlDoc.getelementsbytagname("script")(i).innerhtml
                htmlDoc.getElementsByTagName("script")(i).innerHTML = Replace(htmlDoc.getElementsByTagName("script")(i).innerHTML, "alert", "allert")
            End If
        Next i
    End If
   
    For i = 0 To htmlDoc.all.Length - 1
        If Right(htmlDoc.all(i).tagname, 5) = ("frame") Then
            If htmlDoc.all(i).contentDocument.getElementsByTagName("script").Length <> 0 Then
                'Debug.Print htmlDoc.all(i).Name
                Call retiraAlertJS(htmlDoc.all(i).contentDocument)
            End If
        End If
    Next i
fim:
    If Err.Number <> 0 Then
    End If
End Sub




Public Sub SafeCloseIE()
    'fecha safe o IE
    On Error Resume Next
    ie.Quit
    Set ie = Nothing
    On Error GoTo 0
End Sub



' ----------------------------------------------------------------
' Procedure Name: setPropertiesIES
' Purpose: Mudar propriedades de 1 unico ie ou de todos
' Parameter url (String): Url do nosso internet explorer se branco , modifica de todos
' Parameter AddressBar (Boolean): Barra de endereço Show or not
' Parameter MenuBar (Boolean): Barra de menu Show or not
' Parameter StatusBar (Boolean): Status bar shor or not
' Parameter Toolbar (Boolean): Barra de ferramenta show or not
' Parameter TheatherMode (Boolean): Teather mode é quando aperta f11
' Parameter visible (Boolean):
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Function setPropertiesIES(Optional ByVal url As String = "", Optional ByVal AddressBar As Boolean = True, _
 Optional ByVal MenuBar As Boolean = True, _
 Optional ByVal StatusBar As Boolean = True, _
 Optional ByVal Toolbar As Boolean = True, _
 Optional ByVal TheatherMode As Boolean = False, _
 Optional ByVal visible As Boolean = True)
 
    Dim MyWindows As New SHDocVw.ShellWindows
    Dim window As SHDocVw.InternetExplorer
    For Each window In MyWindows
        'debug.print Window.Name
        If window.Name = "Internet Explorer" Then
            If window.LocationURL = IIf(url = "", window.LocationURL, url) Then
            window.AddressBar = AddressBar
            window.MenuBar = MenuBar
            window.StatusBar = StatusBar
            window.Toolbar = Toolbar
            window.TheaterMode = TheatherMode
            window.visible = visible
            End If
        End If
    Next window
End Function


Public Function returnAllHWND_IEs(Optional ByVal url As String = "") As Variant
'RONAN VICO 02 04 2018
    Dim HWNDs_IE() As Variant
    Dim ieCount As Integer
    Dim MyWindows As New SHDocVw.ShellWindows
    Dim window As SHDocVw.InternetExplorer
    For Each window In MyWindows
        'debug.print Window.Name
        If window.Name = "Internet Explorer" Then
           ieCount = ieCount + 1
           ReDim Preserve HWNDs_IE(ieCount)
           If window.visible = True Then
                HWNDs_IE(ieCount) = window.hwnd
           End If
        End If
    Next window
    returnAllHWND_IEs = HWNDs_IE
End Function

Public Sub ChangeIeVisible(Optional ByVal url As String = "")
    'RONAN VICO 02 04 2018
    Dim MyWindows As Object
    Dim window As InternetExplorer
    
    Set MyWindows = CreateObject("SHDocVw.ShellWindows")
    For Each window In MyWindows
        'debug.print Window.Name
        If window.Name = "Internet Explorer" Then
            If url = "" Then
                window.visible = Not window.visible
            ElseIf Replace(url, "/", "\") = Replace(Replace(window.LocationURL, "%20", " "), "/", "\") Or InStr(1, Replace(Replace(window.LocationURL, "%20", " "), "/", "\"), Mid(Replace(url, "/", "\"), 9)) Then
                window.visible = Not window.visible
            End If
        End If
    Next window
End Sub


' ----------------------------------------------------------------
' Procedure Name: DownloadFileSyncIE
' Purpose: Aceitar o alerta de "Salvar" no internet explorer
' quando clicamos em um link de salvar aparece uma barra na posição inferior do internet explorer
'mandamos um ALT S para salvar ,
' Parameter save_as (String):
' Return Type: String
' Author: Ronan Vico
' Date: 05/03/2019
' ----------------------------------------------------------------
Public Function DownloadFileSyncIE(ByVal save_as As String) As String
'IMPORTANTISSIMO -> É NECESSARIO O IE ESTAR VISIVEL
Dim tentativa As Integer
On Error GoTo trata_erro
inicio:

DownloadFileSyncIE = ""
    Dim contador As Integer: contador = 0
    Dim ieHwnds
    Dim j
    Dim i
    Const dl_key = "HKCU\Software\Microsoft\Internet Explorer\Main\Default Download Directory"
    Static shl As Object
    'Dim waiter As New waiter
    If shl Is Nothing Then
        Set shl = CreateObject("WScript.Shell")
        shl.regWrite "HKCU\Software\Microsoft\Internet Explorer\Main\NotifyDownloadComplete", "no", "REG_SZ"
        shl.regWrite "HKCU\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_FILEDOWNLOAD\iexplore.exe", 0, "REG_DWORD"
    End If

    Dim ie_hwnd, frm_hwnd, n&, folder_bak$, file_name$, tmp_dir$
    
    Call CriaPasta(Replace(save_as, Split(save_as, "\")(UBound(Split(save_as, "\"))), ""))
    ' wait for the download dialogue (IEFrame/Frame Notification Bar/DirectUIHWND)
    'Debug.Print Now
    Do
        'Recebe todas janelas dos explrer abertos
        ieHwnds = returnAllHWND_IEs
        'Verifica se conseguiu pegar janelas , as vez da um erro se tiver com pal no Windows Explorer
        If Len(Join(ieHwnds)) = 0 Then
            'tenta pega a janela mesmo se o window sexplorer tiver pal
            ie_hwnd = FindWindowExA(0, 0, "IEFrame", vbNullString)
            'Se não pegar tambem então da erro
            If ie_hwnd = 0 Then Err.Raise 23, , "Não há explorer aberto !!"
        End If
        
        If Len(Join(ieHwnds)) <> 0 Then
            For n = 0 To UBound(ieHwnds)
                ie_hwnd = ieHwnds(n)
                frm_hwnd = FindWindowExA(ie_hwnd, 0, "Frame Notification Bar", vbNullString)
                If frm_hwnd Then
                    If GetWindowLongA(frm_hwnd, -16) And &H10000000 Then Exit Do    ' If visible
                End If
            Next n
        Else
            frm_hwnd = FindWindowExA(ie_hwnd, 0, "Frame Notification Bar", vbNullString)
            If frm_hwnd Then
               If GetWindowLongA(frm_hwnd, -16) And &H10000000 Then Exit Do    ' If visible
            End If
        End If
        If contador = 400 Then Err.Raise 5, , "Não foi encontrado janela de download"
        contador = contador + 1
        DoEvents
    Loop
    'Debug.Print Now

    ' save the download folder path and create a temporary folder
    tmp_dir = Environ$("TEMP") & "\dl-ie-4f521"
    On Error Resume Next
    folder_bak = shl.regread(dl_key)
    MkDir tmp_dir
    Kill tmp_dir & "\*"
    On Error GoTo 0

    ' set the download folder in the registry
    shl.regWrite dl_key, tmp_dir, "REG_SZ"

    ' send the shortcut for Save (Alt + S)

    ' wait for the file to be downloaded
    j = 0
    Do While InStr(1, file_name, ".partial") Or Len(file_name) = 0
        If j = 100 Then Err.Raise 1234, , "PROBLEMA AO ACEITAR ALERTA DE DOWNLOAD DE ARQUIVO!"
        'call Wait 2000
        PostMessageA frm_hwnd, &H104&, &H12, &H20000001  'WM_SYSKEYDOWN, VK_MENU
        PostMessageA frm_hwnd, &H104&, &H53, &H20000001  'WM_SYSKEYDOWN, S
        PostMessageA frm_hwnd, &H105&, &H53, &HC0000001  'WM_SYSKEYUP, S
        PostMessageA frm_hwnd, &H101&, &H12, &HC0000001  'WM_KEYUP, VK_MENU
        
        
        PostMessageA frm_hwnd, &H104&, &H12, &H20000001  'WM_SYSKEYDOWN, VK_MENU
        PostMessageA frm_hwnd, &H104&, &H53, &H20000001  'WM_SYSKEYDOWN, S
        wait (500)
        PostMessageA frm_hwnd, &H105&, &H53, &HC0000001  'WM_SYSKEYUP, S
        PostMessageA frm_hwnd, &H101&, &H12, &HC0000001  'WM_KEYUP, VK_MENU
        
        
        'call Wait 400
        For i = 0 To 2000
            DoEvents
        Next i
        j = j + 1
        file_name = VBA.Dir(tmp_dir & "\*")
    Loop

    If folder_bak = Empty Then
        shl.RegDelete dl_key
    Else
        shl.regWrite dl_key, folder_bak, "REG_SZ"
    End If

    If Right(save_as, 1) = "\" Then
        save_as = save_as & file_name
    ElseIf InStr(1, Right(save_as, 5), ".") = 0 Then
        save_as = save_as & "." & Split(file_name, ".")(1)
    End If

    save_as = Replace(save_as, Chr(10), "")
    
    ' delete existing file
    If Len(VBA.Dir$(save_as, vbNormal)) Then
        Kill save_as
    ElseIf Len(VBA.Dir$(save_as, vbDirectory)) Then
        'save_as = save_as & "\" & file_name
        RmDir (save_as)
        'If Len(VBA.Dir$(save_as, vbNormal)) Then Kill save_as
    End If
    ' move the file to the provided path
    On Error GoTo Err_handler
    Name tmp_dir & "\" & file_name As save_as
    DownloadFileSyncIE = save_as
    Exit Function
trata_erro:
    If tentativa < 3 Then
        tentativa = tentativa + 1
        Resume inicio
        Resume
    End If
    Err.Raise Err.Number, , Err.Description
Err_handler:
    save_as = ""
    
End Function


' ----------------------------------------------------------------
' Procedure Name: CriaPasta
' Purpose:  Cria uma pasta , usando comandos do shell
' Parameter vcPasta (String): A pasta q vc quer criar (caminho)
' Parameter waitDir (Boolean): se espera ou não a pasta ser criada , util quando usadas em servidores q dependem de tempo de resposta
' Return Type: Boolean
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Function CriaPasta(vcPasta As String, Optional ByVal waitDir As Boolean = True) As Boolean
On Error GoTo fim
    Dim i As Long
    CriaPasta = False
    vcPasta = vcPasta & IIf(Right(vcPasta, 1) <> "\", "\", "")

    'Para não criar na raiz c:\ d:\ e:\, etc
    If Len(vcPasta) <= 3 Then Exit Function

    If Dir(vcPasta, vbDirectory) = "" Then
        Shell "cmd.exe /c  Mkdir """ & vcPasta & """", vbHide
    End If
    If waitDir Then
        While Dir(vcPasta, vbDirectory) = ""
            DoEvents
            i = i + 1
            If i = 10000 Then GoTo fim
        Wend
    End If
    CriaPasta = True

fim:
End Function


' ----------------------------------------------------------------
' Procedure Name: RecursiveGetElementFunction
' Purpose:'Função que percorre a pagina procurando o elemento pelos parametros
'          'necessario recursiva devido a frames dento de frames.
' Parameter OneAttributeNameOneValueName (Variant): ParamARray onde recebe Um attributo
' e um valor do attributo , por exemplo (20,"type","input") ou então (20,"tagname","div") , ou então (20,"type","input","tagname","div")
' Return Type: IHTMLElement
' Author: Ronan Vico
' Date: 05/06/2018
' ----------------------------------------------------------------
Public Function getElement(ByVal SEGUNDOS_PARA_FICAR_TENTANDO As Long, ParamArray OneAttributeNameOneValueName()) As IHTMLElement
'Pra nao cair em loops infinitos
If SEGUNDOS_PARA_FICAR_TENTANDO > 60 Then SEGUNDOS_PARA_FICAR_TENTANDO = 60
If SEGUNDOS_PARA_FICAR_TENTANDO < 0 Then SEGUNDOS_PARA_FICAR_TENTANDO = 1
On Error GoTo Tratar_Erro
  Dim param As Variant
  Dim elemRetorno As Object
  Dim arrAux    As Variant
  Dim DURAÇÃO   As Date

  DURAÇÃO = VBA.Now()
  

  If UBound(OneAttributeNameOneValueName, 1) = 0 Or UBound(OneAttributeNameOneValueName, 1) Mod 2 <> 1 Then
    Err.Raise 666, "getElement", "Array invalido , Digite 1 attributo e 1 Valor do attributo , exemplo: ""tagname"",""div"",""class"",""myClass"""
  End If
  
  For Each param In OneAttributeNameOneValueName
    If param = "" Then
      Err.Raise 667, "GetElement", "Não pode passar parametros em branco!"
    End If
    param = VBA.LCase(param)
  Next
  
  'Fica tentando N segundos ou até que o elemento ja seja encontrado
  While (VBA.Now() < DURAÇÃO + VBA.TimeSerial(0, 0, SEGUNDOS_PARA_FICAR_TENTANDO)) And (elemRetorno Is Nothing)
    arrAux = OneAttributeNameOneValueName
    Set elemRetorno = RecursiveGetElementFunction(ie.Document, arrAux)
  Wend
  
  Set getElement = elemRetorno
  Exit Function
Tratar_Erro:
  Stop
    Debug.Print Err.Number; Err.Description
  Resume
End Function


' ----------------------------------------------------------------
' Procedure Name: RecursiveGetElementFunction
' Purpose:'Função que percorre a pagina procurando o elemento pelos parametros
'          'necessario recursiva devido a frames dento de frames.
' Parameter doc (Object): Documento HTML , onde iremos percorrer todos os elementos
' Parameter OneAttributeNameOneValueName (Variant): ParamARray onde recebe Um attributo
' e um valor do attributo , por exemplo (20,"type","input") ou então (20,"tagname","div") , ou então (20,"type","input","tagname","div")
' Return Type: IHTMLElement
' Author: Ronan Vico
' Date: 05/06/2018
' ----------------------------------------------------------------
Private Function RecursiveGetElementFunction(doc As Object, ByRef OneAttributeNameOneValueName As Variant) As IHTMLElement
On Error GoTo Tratar_Erro
'ATUALIZADA DIA 06 11 2018 , ERRO NO CONTENTDOCUMENT DO FRAME
'Se você for usar essa função tem q arrumar isso usando a biblioteca html
'Made by RV
    Set RecursiveGetElementFunction = Nothing
    
    Dim e
    Dim ELEMENTS
    Dim eFrame As IHTMLFrameElement
    Dim eiframe As IHTMLIFrameElement
    Dim elemAux As Object
    Dim varAux  As Variant
    Dim Props
    Dim propValue
    Dim ename As String, eclass As String
    Dim i
    Dim j As Long
   ' Dim fr As HTMLFrameElement
    Set ELEMENTS = doc.all

    For Each e In ELEMENTS
    
        If needToCheck(e.tagname) Then
            'Se for um frame , é necessario repassar o ddocumento , por isso é uma recursiva
           If VBA.LCase(Right(e.tagname, 5)) = "frame" Then
                'chama a recursividade , sei que parece ambiguo
                'porém é necessario setar um elemento do tipo FRAMe ou IFRAME corretamente
                'pois se não, não ira funcionar em algumas paginas mais atualizadas onde o Documento não ia até o elemento
                Select Case VBA.LCase(e.tagname)
                    Case "iframe"
                      Set eiframe = e
                      Set RecursiveGetElementFunction = RecursiveGetElementFunction(eiframe.contentDocument, OneAttributeNameOneValueName)
                    Case "frame"
                      Set eFrame = e
                      Set RecursiveGetElementFunction = RecursiveGetElementFunction(eFrame.contentDocument, OneAttributeNameOneValueName)
                    Case Else
                      Stop
                End Select
                If Not RecursiveGetElementFunction Is Nothing Then
                    Exit Function
                End If
            Else
              'Some elements will not have the Attribute
              On Error GoTo prox
                For i = LBound(OneAttributeNameOneValueName, 1) To UBound(OneAttributeNameOneValueName, 1) Step 2
                  Props = VBA.Split(OneAttributeNameOneValueName(i), ".")
                  'Verifica se é um objeto dentro de objeto por exemplo "elem.parentNode.Tagname"
                  'Se for , é necessario "Setar" os objetos em cadeia até chegar na ultima propriedade
                  If (UBound(Props)) > 0 Then
                    Set elemAux = e
                    'Vai setando os objetos de dentro até chegar na propriedade
                    For j = LBound(Props) To UBound(Props) - 1
                        Set elemAux = CallByName(elemAux, Props(j), VbGet)
                    Next j
                    
                    If Not VBA.Trim(VBA.LCase(CallByName(elemAux, Props(j), VbGet))) Like VBA.LCase(OneAttributeNameOneValueName(i + 1)) Then
                        'Condição invalida , vai para o proximo elemento
                        GoTo prox
                    End If
                  ElseIf Not VBA.LCase(CallByName(e, OneAttributeNameOneValueName(i), VbGet)) Like VBA.LCase(OneAttributeNameOneValueName(i + 1)) Then
                    'Condição invalida , vai para o proximo elemento
                    GoTo prox
                  End If
                Next i
                'Apenas para garantir esse if
                If i - 1 = UBound(OneAttributeNameOneValueName, 1) Then
                  Set RecursiveGetElementFunction = e
                  Exit Function
                End If
prox:
                  'Stop
                If Err.Number <> 0 Then
                    On Error GoTo Tratar_Erro
                    Resume prox
                End If
            End If
        End If
        
    Next e

    Set ELEMENTS = Nothing
    Exit Function
Tratar_Erro:
  Stop
    Debug.Print Err.Number; Err.Description
  Resume
End Function


Private Function needToCheck(ByVal tagname As String) As Boolean
'Função criada pra não ficar perdendo tempo em algumas tags
needToCheck = False
    tagname = UCase(tagname)
    Select Case tagname
        Case "HTML"
        Case "LINK"
        Case "STYLE"
        Case "BODY"
        Case "SCRIPT"
        Case "HEAD"
        Case "META"
        Case "TITLE"
        Case "NOSCRIPT"
        Case Else
            needToCheck = True
    End Select
End Function

' ----------------------------------------------------------------
' Procedure Name: SendToWindowSaveAs
' Purpose:'Find input Name: on window "Save as ..." , Send the Text value and press Enter.
'            'I did that trying a lot of things , and that one solved, so sorry for my bad procedure.
'            'BUT I WORKD HARD FOR THIS FUCNTION XD , MY BRAIN IS BURNING
' Parameter save_as (String):
' Return Type: String
' Author: Ronan Vico
' Date: 01/03/2018
' ----------------------------------------------------------------
Public Function SendToWindowSaveAs(ByVal save_as As String, Optional ByVal TextoDoTituloDaJanela As String) As String

  On Error GoTo RIP
    Dim hwnd, AUX_hwnd, AUX2_hwnd, txtLen As Integer, FChar As String, file_name As String
    Dim text2 As String
    Dim i As Long
    
    'Find Save as Window
    hwnd = 0
    For i = 1 To 15
        
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", TextoDoTituloDaJanela)
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Escolher arquivo a carregar")
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Salvar Imagem")
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Salvar")
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Save Image")
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Salvar Saída de impressão como")
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "save output print as")
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Salvar página da Web")
        If hwnd <> 0 Then Exit For
        Call wait(500)
    Next i
    
    If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", vbNullString)
    
    If hwnd Then
        Dim hwndDUIViewWndClassName&, hwndDirect&, hwndFNSink&, hwndComboBox&, hwndEdit&
        '-------------DUIViewWndClassName
        For i = 0 To 30
            hwndDUIViewWndClassName = FindWindowExA(hwnd, 0, "DUIViewWndClassName", vbNullString)
            If hwndDUIViewWndClassName Then Exit For
            Call wait(200)
        Next i: If i = 30 Then Err.Raise 7, "NÃO FOI ENCONTRADO JANELA DE SALVAR"
        '----------------------------------DirectUIHWND
        For i = 0 To 30
            hwndDirect = FindWindowExA(hwndDUIViewWndClassName, 0, "DirectUIHWND", vbNullString)
            If hwndDirect Then Exit For
            Call wait(200)
        Next i: If i = 30 Then Err.Raise 7, "NÃO FOI ENCONTRADO JANELA DE SALVAR"
        '---------------------------------FloatNotifySink
        For i = 0 To 30
            hwndFNSink = FindWindowExA(hwndDirect, 0, "FloatNotifySink", vbNullString)
            If hwndFNSink Then Exit For
            Call wait(200)
        Next i: If i = 30 Then Err.Raise 7, "NÃO FOI ENCONTRADO JANELA DE SALVAR"
        '-----------------------------------ComboBox
        For i = 0 To 30
            hwndComboBox = FindWindowExA(hwndFNSink, 0, "ComboBox", vbNullString)
            If hwndComboBox Then Exit For
            Call wait(200)
        Next i: If i = 30 Then Err.Raise 7, "NÃO FOI ENCONTRADO JANELA DE SALVAR"
        '----------------------------Edit
        For i = 0 To 30
            hwndEdit = FindWindowExA(hwndComboBox, 0, "Edit", vbNullString)
            If hwndEdit Then Exit For
            Call wait(200)
        Next i: If i = 30 Then Err.Raise 7, "NÃO FOI ENCONTRADO JANELA DE SALVAR"
            
            'DUIViewWndClassName -> DirectUIHWND -> FloatNotifySink -> ComboBox -> Edit
        
        'Send Some Char with PostMessage or this function wont work
        SetForegroundWindow (hwndEdit)
        PostMessageA hwndEdit, &H100, Asc("C"), 0    '
        Call wait(250)
        SetForegroundWindow (hwndEdit)
        PostMessageA hwndEdit, &H101, Asc("C"), 0    '
        Call wait(250)
        
        'While is visible "Save As window , try to put foreground and send text + enter
        If VBA.Dir(save_as) <> "" Then Kill save_as
        i = 0
        While FindWindowExA(hwnd, 0, vbNullString, vbNullString) And i < 50
        
            SetForegroundWindow (hwndEdit)
            SendMessageByString hwndEdit, &HC, Len(save_as), save_as
            'send 2 times for make sure that is going to work
            SendMessageByString hwndEdit, &HC, Len(save_as), save_as
            SetForegroundWindow (hwndEdit)
            PostMessageA hwndEdit, &H100, &HD, 0     ' '&h100 = MW_KDOWN , 101 = MW_KUP
            PostMessageA hwndEdit, &H101, &HD, 0     ' ' &HD = ENTER ON ASCII HEX
            i = i + 1
        Wend
        Call wait(5000)
    End If
    SendToWindowSaveAs = save_as
  
    Exit Function
    Resume

RIP:
    Select Case Err.Number
      Case 0
      Case Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SendToWindowSaveAs, line " & Erl & "."
    End Select
End Function





' ----------------------------------------------------------------
' Procedure Name: SendToWindowOpen
' Purpose:'Find input Name: on window "Open" , Send the Text value and press Enter.
'            'I did that trying a lot of things , and that one solved, so sorry for my bad procedure.
'            'BUT I WORKD HARD FOR THIS FUCNTION XD , MY BRAIN IS BURNING
' Parameter text (String):
' Return Type: Boolean
' Author: Ronan Vico
' Date: 23/04/2018
' ----------------------------------------------------------------
Public Function SendToWindowOpen(ByVal text As String) As Boolean
SendToWindowOpen = False
    Dim hwnd As Long, hwnd_ComboBoxEx32 As Long, hwnd_ComboBox As Long, hwnd_Edit As Long, i As Integer, txtLen As Integer, FChar As String
    Dim text2 As String

    'Find Save as Window
    hwnd = 0

    'espera alguns segundos pela janela de abrir
    For i = 0 To 10
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Open")
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Abrir")
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Abrir Como")
        If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Escolher arquivo a carregar")
        If hwnd <> 0 Then Exit For
        
        Call wait(500)
    Next i
    
    'se não encontrou a janela sai da função
    If hwnd = 0 Then
        'MsgBox "JANELA DE ABRIR NÃO ENCONTRADA", vbCritical
        Exit Function
    End If
    
    'If hWnd = 0 Then hWnd = FindWindowExA(0, 0, "#32770", vbNullString)

    hwnd_ComboBoxEx32 = FindWindowExA(hwnd, 0, "ComboBoxEx32", vbNullString)
    hwnd_ComboBox = FindWindowExA(hwnd_ComboBoxEx32, 0, "ComboBox", vbNullString)
    hwnd_Edit = FindWindowExA(hwnd_ComboBox, 0, "Edit", vbNullString)

    'Debug.Print vbNewLine & WindowClass(AUX_hwnd) & vbNewLine & WindowTextString(AUX_hwnd) & vbNewLine

    'Send Some Char with PostMessage or this function wont work

    SetForegroundWindow (hwnd_Edit)
    PostMessageA hwnd_Edit, &H101, Asc("C"), 0    '
    Call wait(250)

    'Send Some Char with PostMessage or this function wont work
    SetForegroundWindow (hwnd_Edit)
    PostMessageA hwnd_Edit, &H100, Asc("C"), 0    '
    Call wait(250)
    SetForegroundWindow (hwnd_Edit)
    PostMessageA hwnd_Edit, &H101, Asc("C"), 0    '
    Call wait(250)
    

    'While is visible "Save As window , try to put foreground and send text + enter
    i = 0
    While FindWindowExA(hwnd_ComboBox, 0, "Edit", vbNullString) And i < 50
        SetForegroundWindow (hwnd_Edit)
        SendMessageByString hwnd_Edit, &HC, Len(text), text
        SendMessageByString hwnd_Edit, &HC, Len(text), text
        Call wait(300)
        SetForegroundWindow (hwnd_Edit)
        PostMessageA hwnd_Edit, &H100, &HD, 0     ' '&h100 = MW_KDOWN , 101 = MW_KUP
        PostMessageA hwnd_Edit, &H101, &HD, 0     ' ' &HD = ENTER ON ASCII HEX
        i = i + 1
    Wend

    SendToWindowOpen = True
End Function


' ----------------------------------------------------------------
' Procedure Name: SendEnterToSaveOrOpenWindow
' Purpose: FEITA POR RONAN VICO , ENVIANDO ENTER PARA UMA JANELA DE SALVAR ARQUIVO COMO.
'           'DEU TRABALHO PARA FAZER ESTA FUNÇÃO APESAR DE SER PEQUENA.
' Author: Ronan Vico
' Date: 28/02/2018
' ----------------------------------------------------------------
Public Sub SendEnterToSaveOrOpenWindow()
    Dim hwnd
    If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Escolher arquivo a carregar")
    If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Salvar Imagem")
    If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Salvar")
    If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Save Image")
    If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Imprimir")
    If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Print")
    If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "Salvar Saída de impressão como")
    If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", "save output print as")
    If hwnd = 0 Then hwnd = FindWindowExA(0, 0, "#32770", vbNullString)
    'ENVIA A JANELA MY_HWND , WM_KEY PRA BAIXO , ENTER EM HEXADECIMAL
    PostMessageA hwnd, &H100, &HD, 0     '
    'ENVIA A JANELA MY_HWND , WM_KEY PARA CIMA , ENTER EM HEXADECIMAL
    PostMessageA hwnd, &H101, &HD, 0     '
End Sub


Public Sub MouseRightDown()
    'Clica com lado direito na posição que  o cursor do mouse estiver
    mouse_event &H8, 0, 0, 0, 0 'Mouse  right up
    mouse_event &H10, 0, 0, 0, 0 'Mouse right down
End Sub

Public Sub MouseLeftDown()
    'Clica com lado esquerdo  na posição que  o cursor do mouse estiver
    mouse_event &H2, 0, 0, 0, 0 ' mouse left up
    mouse_event &H4, 0, 0, 0, 0 ' mouse left down
End Sub

Public Sub MoverMousePara(point As POINTAPI)
    Call SetCursorPos(point.X_Pos, point.Y_Pos)
End Sub

Public Sub PegarPosicoesDoMouse(ByRef point As POINTAPI)
    Call GetCursorPos(point)
End Sub

' ----------------------------------------------------------------
' Procedure Name: getWindowClass
' Purpose: RETORNA O NOME DA WINDOW FRM_HWND = O LONG DA SUA WINDOW
' Parameter hwnd (Long):
' Return Type: String
' Author: Ronan Vico
' Date: 28/02/2018
' ----------------------------------------------------------------
Private Function WindowTextString(ByVal hwnd As Long) As String
    Dim lenTxt As Long, retText As String
    lenTxt = GetWindowTextLength(hwnd) + 1
    retText = String$(lenTxt, " ")
    GetWindowText hwnd, retText, lenTxt
    WindowTextString = Mid(retText, 1, lenTxt - 1)
End Function

' ----------------------------------------------------------------
' Procedure Name: getWindowClass
' Purpose: RETORNA O NOME DA CLASSE EM #NUMEROS EXEMPLO #12312312
' Parameter hwnd (Long):
' Return Type: String
' Author: Ronan Vico
' Date: 28/02/2018
' ----------------------------------------------------------------
Private Function getWindowClass(ByVal hwnd As Long) As String
    Dim s As String
    s = Space$(255)
    Call GetClassName(hwnd, s, Len(s))
    If InStr(1, s, vbNullChar) Then
        s = left$(s, InStr(1, s, vbNullChar) - 1)
    End If
    getWindowClass = s
End Function

' ----------------------------------------------------------------
' Procedure Name: ExecScriptAssync
' Purpose: Executa um Javascripto , porém após alguns segundos ,
'          o motivo disso é evitar que o objeto internet explorer fique travado após
'           algum comando.
' Parameter js (String):
' Parameter seconds (Long):
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Sub ExecScriptAssync(ByVal js As String, ByVal seconds As Long)
    seconds = seconds * 1000
    If VBA.InStr(js, VBA.Chr(34)) <> 0 Then
        Err.Raise 123456, "ExecScriptAssync()", "Não é possivel um JS com Aspas Duplas , use ' (aspas simples no lugar)"
    End If
    Call execScript("setTimeout(""" & js & """, " & seconds & ")")
End Sub


' ----------------------------------------------------------------
' Procedure Name: TableToRange
' Purpose: Pega um ELEMENTO html do tipo <table> e joga em um range do excel.
' Parameter TABLE (IHTMLTable): Elemento HTML
' Parameter rg (Range): Range do excel
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Sub TableToRange(ByRef TABLE As IHTMLTable, rg As Range)
    Dim r As Long, c As Long
    For r = 0 To TABLE.Rows.Length - 1
        For c = 0 To TABLE.Rows(r).Cells.Length - 1
            rg.Offset(r, c).Value = TABLE.Rows(r).Cells(c).InnerText
        Next
    Next r
End Sub


'Deprecated not work anymore
'Public Sub MoverMouseParaElemento(ByVal elem As IHTMLElement)
'On Error GoTo tratarErro
'    Dim top As Long
'    Dim left As Long
'    Dim elemAux As IHTMLElement
'
'
'    left = elem.Width * -1
'    top = elem.Height * -1
'    Set elemAux = elem
'    On Error GoTo fim
'    While Not elemAux.ParentNode Is Nothing
'        left = left + elemAux.offsetLeft
'        top = top + elemAux.offsetTop
'        Set elemAux = elemAux.ParentNode
'    Wend
'
'fim:
'    On Error GoTo tratarErro
'    left = left + Me.ie.left
'    top = top + Me.ie.top
'    Call SetCursorPos(left, top)
'    Exit Sub
'tratarErro:
'    Stop
'    Resume
'End Sub

' ----------------------------------------------------------------
' Procedure Name: quantos_ies_abertos
' Purpose: Conta quantos IEs estão abertos Visiveis ou não
' Return Type: Long
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Function quantos_ies_abertos() As Long
Dim MyWindows As Object
Dim window As Variant
Set MyWindows = CreateObject("SHELL.APPLICATION").Windows
For Each window In MyWindows
    'Debug.Print Window.Name
    If window.Name = "Internet Explorer" Then
       quantos_ies_abertos = quantos_ies_abertos + 1
    End If
Next window
End Function


' ----------------------------------------------------------------
' Procedure Name: TrazerIeParaFrente
' Purpose: Traz o Internet explorer  para frente
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Sub TrazerIeParaFrente()
    SetForegroundWindow Me.ie.hwnd
End Sub

' ----------------------------------------------------------------
' Procedure Name: SetPrintPDF
' Purpose: Altera a impressora padrão
' Parameter Printer (String): Nome da impressora
' Author: Ronan Vico
' ----------------------------------------------------------------
Public Sub SetPrintPDF(Optional ByVal Printer As String = "Microsoft Print to PDF")
    Dim WSHNetwork As Object
    Set WSHNetwork = CreateObject("WScript.Network")
    WSHNetwork.SetDefaultPrinter Printer
    VBA.Shell ("RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n" & Chr(34) & Printer & Chr(34))
End Sub

