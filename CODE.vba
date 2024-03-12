Public SapGuiAuto As Variant
Public SAPApp As Variant
Public SAPCon As Variant
Public session As Variant
Public Connection As Variant
Public WScript As Variant

Sub ConectarSAP()

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)
    If Not IsObject(Application) Then
        Set SapGuiAuto = GetObject("SAPGUI")
    End If
    If Not IsObject(session) Then
        Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject Application, "on"
    End If
End Sub

Sub Main()
    ConectarSAP
    EntrarTela
End Sub

Sub EntrarTela()
    For Linha = 2 To 21
        Ordem = ActiveSheet.Cells(Linha, 2).Value
        EntrarTelaSAP (Ordem)
        
        
        Quantidade = 0
        
        For LinhaSAP = 0 To 25:
            
            ComponenteDenominacao = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MATXT[2," & LinhaSAP & "]").Text
            ComponenteNumero = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & LinhaSAP & "]").Text
            
            
            If InStr(1, ComponenteDenominacao, "ANEL") <> 0 Then
                ComponenteQuantia = session.findById("/app/con[0]/ses[0]/wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3," & LinhaSAP & "]").Text
                ComponenteQuantia = CDec(ComponenteQuantia)
                
                Quantidade = Quantidade + ComponenteQuantia
                
                Debug.Print "CHAPA ROTORE QUANTIA:"
                Debug.Print (ComponenteQuantia)
                
            End If
        
        Next LinhaSAP
        
        ActiveSheet.Cells(Linha, 3).Value = Quantidade
        
    Next Linha
End Sub

Sub EntrarTelaSAP(Ordem)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nco03"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = Ordem
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[6]").press
End Sub
