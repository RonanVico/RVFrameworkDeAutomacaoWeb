Attribute VB_Name = "BAS_RIBBON"
Option Explicit



Public Sub CallNavigateMidia(control As IRibbonControl)
'    Stop
    Dim url
    Select Case LCase(control.id)
        Case LCase("Linkedin")
            url = "https://br.linkedin.com/in/ronan-vico"
        Case LCase("GitHub")
            url = "https://github.com/RonanVico"
    End Select
    Call ActiveWorkbook.FollowHyperlink(url)
End Sub

Public Sub CallFormDados(control As IRibbonControl)
    Form_Contato.Show
End Sub
