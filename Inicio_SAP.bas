Attribute VB_Name = "Inicio_SAP"
Sub SAP_Start()
    
    Dim Appl As Object, Connection As Object, session As Object, WshShell As Object, SapGui As Object
    Dim user As String, Hoy As Date, Mañana As Date, Mes As Date, Fecha_SAP_Hoy As String, Fecha_SAP_Mañana As String, Fecha_SAP_Mes As String
    Dim psw As String
    
    'Fechas
    Hoy = Date
    Mañana = Date + 1
    Mes = Date + 30
        
    ' Formato SAP a las fechas
    Fecha_SAP_Hoy = Format(Hoy, "DD.MM.YYYY")
    Fecha_SAP_Mañana = Format(Mañana, "DD.MM.YYYY")
    Fecha_SAP_Mes = Format(Mes, "DD.MM.YYYY")
        
    ' Definición de entradas a SAP
    user = "USER"
    psw = "PASSWORD"
    
    ' Directorio de ejecución de SAP
    Shell "C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe", 4
    
    ' Espera hasta que SAP Logon esté activo
    Set WshShell = CreateObject("WScript.Shell")
    Do Until WshShell.AppActivate("SAP Logon ")
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    Set WshShell = Nothing

    ' Conecta con SAP
    Set SapGui = GetObject("SAPGUI")
    Set Appl = SapGui.GetScriptingEngine
    Set Connection = Appl.OpenConnection("R/3 - Productivo", True)
    Set session = Connection.Children(0)
    
    'Datos de ingreso a SAP
    session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "400"
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = user
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = psw
    session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "ES"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    
    'Creación de sesiones con las transacciones necesarias
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/oZSD446"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/oVA03"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/oZSD013"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/oZSD013"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/oVL03N"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    
    'Ingreso a la variante de la transacción ZSD013_G01
    session.StartTransaction ("ZSD013_G01")
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LOROJAS"
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[0]/usr/ctxtP_DPLBG").Text = Fecha_SAP_Hoy
    session.findById("wnd[0]").sendVKey 8
    
    'Ingreso a la variante de la transacción ZSD446
    Set session = Connection.Children(1)
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LOROJAS"
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]").sendVKey 2
    session.findById("wnd[0]/usr/ctxtP_DPLBG").Text = Fecha_SAP_Mes
    session.findById("wnd[0]").sendVKey 8

    'Ingreso a la variante de la transacción ZSD013 Ensacado
    Set session = Connection.Children(3)
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LOROJAS"
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
    session.findById("wnd[1]").sendVKey 2
    session.findById("wnd[0]/usr/ctxtS_DPLBG-LOW").Text = Fecha_SAP_Hoy
    session.findById("wnd[0]/usr/ctxtS_DPLBG-HIGH").Text = Fecha_SAP_Mañana
    session.findById("wnd[0]").sendVKey 8

    'Ingreso a la variante de la transacción ZSD013 Granel
    Set session = Connection.Children(4)
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LOROJAS"
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
    session.findById("wnd[1]").sendVKey 2
    session.findById("wnd[0]/usr/ctxtS_MFRGR-LOW").Text = "02"
    session.findById("wnd[0]/usr/ctxtS_MFRGR-HIGH").Text = "03"
    session.findById("wnd[0]/usr/ctxtS_DPLBG-LOW").Text = Fecha_SAP_Hoy
    session.findById("wnd[0]/usr/ctxtS_DPLBG-HIGH").Text = Fecha_SAP_Mañana
    session.findById("wnd[0]").sendVKey 8
       
End Sub
