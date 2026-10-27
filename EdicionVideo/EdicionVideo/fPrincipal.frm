VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Begin VB.Form fPrincipal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   5295
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   6855
      _cx             =   12091
      _cy             =   9340
      FlashVars       =   ""
      Movie           =   "flashA.swf"
      Src             =   "flashA.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
End
Attribute VB_Name = "fPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ruta As Collection
Dim texto As Collection
Dim valorV As Collection
Dim edicion As Collection

Dim peliculaVideo As Collection
Dim peliculaAudio As Collection
Dim rutasalida As String
Dim nombre As String
Dim fondo As String
Dim tiempoLimite As String
Dim tiempoReset As String
Dim tiempoEspera As String

Dim idiomas As Collection
Dim animaciones As Collection

Private Type t_mailConfig
    smtp As String
    smtpport As Integer
    user As String
    pass As String
    
    to As String
    from As String
    subject As String
    
    enviar As Boolean
End Type
Dim mailConfig As t_mailConfig
    
    
Private Sub Flash_FlashCall(ByVal request As String)
    Dim objParser As MSXML.DOMDocument
    Dim x As MSXML.IXMLDOMNodeList
    Set objParser = New MSXML.DOMDocument
    objParser.async = "false"
    objParser.loadXML request
    
    Dim parametro As Collection
    Set parametro = New Collection
    
    Set x = objParser.getElementsByTagName("invoke")
    funcion = x(0).Attributes.getNamedItem("name").nodeValue
    Set x = objParser.getElementsByTagName("arguments")
    
    For i = 0 To x.length - 1
       If x(i).hasChildNodes Then
        parametro.Add x(i).childNodes(0).Text
       End If
    Next i
    
    If (funcion = "rutavideo") Then
        valor = CInt(parametro.Item(1))
        If (valor <= ruta.Count) Then
            t = ruta.Item(valor)
            Flash.SetReturnValue t
        End If
    End If
    
    If (funcion = "textovideo") Then
        valor = CInt(parametro.Item(1))
        If (valor <= texto.Count) Then
            t = texto.Item(valor)
            If (Not IsNull(t)) Then
                Flash.SetReturnValue t
            End If
        End If
    End If
    
    If (funcion = "valorvideo") Then
        valor = CInt(parametro.Item(1))
        If (valor <= valorV.Count) Then
            t = valorV.Item(valor)
            If (Not IsNull(t)) Then
                Flash.SetReturnValue t
            End If
        End If
    End If
    
    If (funcion = "rutaedicion") Then
        valor = CInt(parametro.Item(1))
        If (valor <= ruta.Count) Then
            t = edicion.Item(valor)
            Flash.SetReturnValue t
        End If
    End If
    
    If (funcion = "tiempoLimite") Then
        Flash.SetReturnValue tiempoLimite
    End If
    
    If (funcion = "tiempoEspera") Then
        Flash.SetReturnValue tiempoEspera
    End If
    
    If (funcion = "tiempoReset") Then
        Flash.SetReturnValue tiempoReset
    End If
    
    If (funcion = "fondo") Then
        Flash.SetReturnValue fondo
    End If
    
    If (funcion = "idioma") Then
        Dim tempIdioma As cIdioma
        Flash.SetReturnValue ""
        valor = CStr(parametro.Item(1))
        For i = 1 To idiomas.Count
            Set tempIdioma = idiomas.Item(i)
            If tempIdioma.id = valor Then
                Flash.SetReturnValue tempIdioma.ruta
            End If
        Next i
    End If

    If (funcion = "animacion") Then
        Dim tempAnim As cIdioma
        Flash.SetReturnValue ""
        valor = CStr(parametro.Item(1))
        For i = 1 To idiomas.Count
            Set tempAnim = animaciones.Item(i)
            If tempAnim.id = valor Then
                Flash.SetReturnValue tempAnim.ruta
            End If
        Next i
    End If

    
    If (funcion = "vaciarpelicula") Then
        Set peliculaVideo = New Collection
        Set peliculaAudio = New Collection
        rutasalida = ""
        nombre = ""
    End If
    
    If (funcion = "rutasalida") Then
        valor = CStr(parametro.Item(1))
        rutasalida = valor
    End If
    
    If (funcion = "nombre") Then
        valor = CStr(parametro.Item(1))
        If valor = "null" Then valor = "("
        nombre = valor
    End If
    
    If (funcion = "añadirvideo") Then
        valor = CStr(parametro.Item(1))
        peliculaVideo.Add (App.Path & "\" & valor)
    End If
    
    If (funcion = "añadiraudio") Then
        valor = CStr(parametro.Item(1))
        peliculaAudio.Add (App.Path & "\" & valor)
    End If
    

End Sub

Private Sub Flash_FSCommand(ByVal command As String, ByVal args As String)
    If command = "alert" Then MsgBox args
    If command = "ESC" Then End
    If command = "crearvideo" Then CrearVideo rutasalida, nombre
End Sub


Private Sub Form_Load()
    Dim objParser As MSXML.DOMDocument
    Dim x As MSXML.IXMLDOMNodeList
    Dim n As MSXML.IXMLDOMNode
    Set objParser = New MSXML.DOMDocument
    objParser.async = "false"
    objParser.Load "ui.xml"
    
    Set ruta = New Collection
    Set texto = New Collection
    Set edicion = New Collection
    Set valorV = New Collection
    Set idiomas = New Collection
    Set animaciones = New Collection
    
    Set x = objParser.getElementsByTagName("objeto")
    For i = 0 To x.length - 1
        t = (x(i).Attributes.getNamedItem("ruta").nodeValue)
        ruta.Add ("<string>" & t & "</string>")
        Set n = x(i).Attributes.getNamedItem("texto")
        If (n Is Nothing) Then
            texto.Add (Null)
        Else
            t = n.nodeValue
            texto.Add ("<string>" & t & "</string>")
        End If
        t = (x(i).Attributes.getNamedItem("edicion").nodeValue)
        edicion.Add ("<string>" & t & "</string>")
        t = (x(i).Attributes.getNamedItem("valor").nodeValue)
        valorV.Add ("<string>" & t & "</string>")
    Next i
    
    Set x = objParser.getElementsByTagName("fondo")
    For i = 0 To x.length - 1
        Set n = x(i).Attributes.getNamedItem("ruta")
        If Not (n Is Nothing) Then
            t = n.nodeValue
            fondo = "<string>" & t & "</string>"
        End If
    Next i
    
    Set x = objParser.getElementsByTagName("tiempoEspera")
    For i = 0 To x.length - 1
        Set n = x(i).Attributes.getNamedItem("valor")
        If Not (n Is Nothing) Then
            t = n.nodeValue
            tiempoEspera = "<string>" & t & "</string>"
        End If
    Next i
    
    Set x = objParser.getElementsByTagName("tiempoLimite")
    For i = 0 To x.length - 1
        Set n = x(i).Attributes.getNamedItem("valor")
        If Not (n Is Nothing) Then
            t = n.nodeValue
            tiempoLimite = "<string>" & t & "</string>"
        End If
    Next i
    
    Set x = objParser.getElementsByTagName("tiempoReset")
    For i = 0 To x.length - 1
        Set n = x(i).Attributes.getNamedItem("valor")
        If Not (n Is Nothing) Then
            t = n.nodeValue
            tiempoReset = "<string>" & t & "</string>"
        End If
    Next i
    
    Set x = objParser.getElementsByTagName("texto")
    For i = 0 To x.length - 1
        Dim tempIdioma As New cIdioma
        Set n = x(i).Attributes.getNamedItem("idioma")
        If Not (n Is Nothing) Then
            t = n.nodeValue
            tempIdioma.id = t
            Set n = x(i).Attributes.getNamedItem("ruta")
            If Not (n Is Nothing) Then
                t = n.nodeValue
                tempIdioma.ruta = "<string>" & t & "</string>"
                idiomas.Add tempIdioma
            End If
        End If
        Set tempIdioma = Nothing
    Next i
    
    Set x = objParser.getElementsByTagName("animacion")
    For i = 0 To x.length - 1
        Dim tempAnim As New cIdioma
        Set n = x(i).Attributes.getNamedItem("idioma")
        If Not (n Is Nothing) Then
            t = n.nodeValue
            tempAnim.id = t
            Set n = x(i).Attributes.getNamedItem("ruta")
            If Not (n Is Nothing) Then
                t = n.nodeValue
                tempAnim.ruta = "<string>" & t & "</string>"
                animaciones.Add tempAnim
            End If
        End If
        Set tempIdioma = Nothing
    Next i
    
    'leer configuracion mail
    Set x = objParser.getElementsByTagName("mail")
    For i = 0 To x.length - 1
        Dim todoOk As Boolean
        todoOk = True
        mailConfig.enviar = False
        Set n = x(i).Attributes.getNamedItem("smtp")
        If ((Not (n Is Nothing)) And todoOk) Then
            mailConfig.smtp = n.nodeValue
        Else
            todoOk = False
        End If
        Set n = x(i).Attributes.getNamedItem("smtpport")
        If ((Not (n Is Nothing)) And todoOk) Then
            mailConfig.smtpport = n.nodeValue
        Else
            todoOk = False
        End If
        Set n = x(i).Attributes.getNamedItem("user")
        If ((Not (n Is Nothing)) And todoOk) Then
            mailConfig.user = n.nodeValue
        Else
            todoOk = False
        End If
        Set n = x(i).Attributes.getNamedItem("pass")
        If ((Not (n Is Nothing)) And todoOk) Then
            mailConfig.pass = n.nodeValue
        Else
            todoOk = False
        End If
        Set n = x(i).Attributes.getNamedItem("to")
        If ((Not (n Is Nothing)) And todoOk) Then
            mailConfig.to = n.nodeValue
        Else
            todoOk = False
        End If
        Set n = x(i).Attributes.getNamedItem("from")
        If ((Not (n Is Nothing)) And todoOk) Then
            mailConfig.from = n.nodeValue
        Else
            todoOk = False
        End If
        Set n = x(i).Attributes.getNamedItem("subject")
        If ((Not (n Is Nothing)) And todoOk) Then
            mailConfig.subject = n.nodeValue
        Else
            todoOk = False
        End If
        If todoOk Then mailConfig.enviar = True
    Next i
    
    Me.Flash.Movie = App.Path + "\ui.swf"
    
End Sub

Private Sub Form_Resize()
    With Flash
    .top = 0
    .left = 0
    .Width = Me.Width
    .Height = Me.Height
    End With
End Sub

Private Sub CrearVideo(ByVal mail As String, nombre As String)
    Dim fileOut As String
    fileOut = App.Path & "\output\" & nombre & "(" & mail & ").avi"
    Dim fso As New FileSystemObject
    If fso.FileExists(fileOut) Then
        fso.DeleteFile fileOut, True
    End If
'    MsgBox fileOut
'    For Each a In peliculaVideo
'        MsgBox a
'    Next
    '// Recibir Parametros
'    Dim szFileOut As String
'    szFileOut = App.Path & "/output/output.avi"
'
'    Dim szFileIn As Collection
'    Set szFileIn = New Collection
'    szFileIn.Add (App.Path & "/video/avi/video1.avi")
'    szFileIn.Add (App.Path & "/video/avi/video2.avi")
    Refresh
    DoEvents
    resultado = JuntarVideos(peliculaVideo, peliculaAudio, fileOut)
    '    MsgBox "video juntos"

' notificar al flash... cuando se complete el envio

    ' ENVIAR POR EMAIL
    If mailConfig.enviar Then
        On Error GoTo errorout
        Dim objMessage As CDO.Message
        Set objMessage = CreateObject("cdo.message")
            
        '// Mail para un receptor comun //
        'objMessage.to = mailConfig.to
        'objMessage.subject = mailConfig.subject
        'objMessage.from = mailConfig.from
        'objMessage.TextBody = nombre
        
        '// Mail directo al usuario. //
        objMessage.to = mail
        objMessage.subject = mailConfig.subject
        objMessage.from = mailConfig.from
        objMessage.TextBody = "Este es tu video editado 32ª America's Cup." + Chr(13) + "This is your edited 32nd America's Cup video."
        
        objMessage.AddAttachment ("file:\\" & fileOut)
            
            'provide the configuration information for the remote SMTP server.
              objMessage.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            'Name or IP of Remote SMTP Server
              objMessage.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mailConfig.smtp
            'Type of authentication, NONE, Basic (Base64 encoded), NTLM
              objMessage.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
            'Your UserID on the SMTP server
              objMessage.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/sendusername") = mailConfig.user
            'Your password on the SMTP server
              objMessage.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = mailConfig.pass
            'Server port (typically 25)
              objMessage.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = mailConfig.smtpport
            'Use SSL for the connection (False or True)
              objMessage.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
            'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
              objMessage.Configuration.Fields.Item _
                ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
              objMessage.Configuration.Fields.Update
            'End remote SMTP server configuration
            
        objMessage.send
        ' mensaje enviandose... (espero)
        resultado = AVIERR_OK
        GoTo errorbien
errorout:
        'MsgBox Err.Number & " _ " & Err.Description
        resultado = -1
errorbien:
    End If

    ' notificar al flash... todo ok
    If (resultado = AVIERR_OK) Then
        Flash.TGotoLabel "_root", "envioOK"
    Else
        Flash.TGotoLabel "_root", "envioError"
        'MsgBox "send error: " & resultado
    End If
    
End Sub
Public Sub InformarCrearVideo(porciento As Long)
    Refresh
    DoEvents
    Flash.SetVariable "_root.porcentaje", CStr(porciento)
End Sub
