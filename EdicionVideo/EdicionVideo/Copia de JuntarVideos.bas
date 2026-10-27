Attribute VB_Name = "Module1"
'Option Explicit
'DefLng A-Z

Private Sub JuntarVideos()
  
    Dim hHeap As Long
    hHeap = GetProcessHeap
    '//pFrameformat = HeapAlloc(hHeap, 0, frameformatsize)     // asignar memoria
    
    Call AVIFileInit  '// opens AVIFile library
    
    Dim szFile, szFileOut As String
    Dim res As Long
   
    szFile = App.Path & "/source/v1.avi"
    szFileOut = App.Path & "/target/output.avi"
        
    '// abrir archivo
    Dim pAVIFile As Long  'pointer to AVI File (PAVIFILE handle)
    res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
    If res <> AVIERR_OK Then GoTo ErrorOut
    MsgBox "fichero abierto"
    
    '// leer stream
    Dim pAVIStream As Long  'pointer to AVI stream interface (PAVISTREAM handle)
    'Get the first available video stream (PAVISTREAM)
    res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    If res <> AVIERR_OK Then GoTo ErrorOut
    MsgBox "stream abierto"
    
        
    '// avi info
    Dim fileInfo As AVI_FILE_INFO  'file info struct
    'get file info struct (UDT)
    res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut
    'print file info to Debug Window
    Call DebugPrintAVIFileInfo(fileInfo)
    MsgBox "avi info"

    '// stream info
    Dim streamInfo As AVI_STREAM_INFO 'stream info struct
    'get stream info struct (UDT)
    res = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut
    'print stream info to Debug Window
    Call DebugPrintAVIStreamInfo(streamInfo)
    MsgBox "stream info"
        
    '// informacion del stream
    Dim numFrames As Long 'number of frames in video stream
    Dim firstFrame As Long 'position of the first video frame
    'get the starting position of the stream (some streams may not start simultaneously)
    firstFrame = AVIStreamStart(pAVIStream)
    If firstFrame = -1 Then GoTo ErrorOut 'this function returns -1 on error
    'get the length of video stream in frames
    numFrames = AVIStreamLength(pAVIStream)
    If numFrames = -1 Then GoTo ErrorOut ' this function returns -1 on error
    MsgBox "first: " & firstFrame & " - num: " & numFrames
        
        
    '// abrir frame
    Dim dib As cDIB
    Dim pGetFrameObj As Long 'pointer to GetFrame interface
    Dim pDIB As Long 'pointer to packed DIB in memory
    Dim bih As BITMAPINFOHEADER 'infoheader to pass to GetFrame functions
    Dim i As Long
    'set bih attributes which we want GetFrame functions to return
    With bih
        .biBitCount = 24
        .biClrImportant = 0
        .biClrUsed = 0
        .biCompression = BI_RGB
        .biHeight = streamInfo.rcFrame.bottom - streamInfo.rcFrame.top
        .biPlanes = 1
        .biSize = 40
        .biWidth = streamInfo.rcFrame.right - streamInfo.rcFrame.left
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        'calculate total size of RGBQUAD scanlines (DWORD aligned)
        .biSizeImage = (((.biWidth * 3) + 3) And &HFFFC) * .biHeight
    End With
    'init AVISTreamGetFrame* functions and create GETFRAME object
    pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, bih) 'force function to return 24bit DIBS
    MsgBox "streamGetframes abierto"
    'create a DIB class to load the frames into
    Set dib = New cDIB
        pDIB = AVIStreamGetFrame(pGetFrameObj, firstFrame) 'returns "packed DIB"
        If dib.CreateFromPackedDIBPointer(pDIB) Then
            'Call dib.WriteToFile(App.Path & "\bmps\" & i & ".bmp")
            'txtStatus = "Bitmap " & i + 1 & " of " & numFrames & " written to app folder"
            'txtStatus.Refresh
        Else
            
        End If
    'Set dib = Nothing
    MsgBox "dib leido primer frame"
    
    MsgBox "todo listo para leer"
    
        
    '// crear un nuevo stream
    '// abrir archivo de salida
    Dim pfileout As Long
    Dim pAVIStreamOut2 As Long
    Dim pAVIStreamCompressed As Long
    '     Open the file for writing
    res = AVIFileOpen(pfileout, szFileOut, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo ErrorOut
    MsgBox "nuevo avi abierto"
    ' Fill in the header for the video stream
    Dim strhdr As AVI_STREAM_INFO
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)          '// stream type video
        .fccHandler = 0&                                   '// default AVI handler
        .dwScale = 1
        .dwRate = 25                                       '// fps
        .dwSuggestedBufferSize = dib.SizeImage             '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, dib.Width, dib.Height) '// rectangle for stream
    End With
    MsgBox "stream info relleno del bitmap"
    ' And create the stream
    res = AVIFileCreateStream(pfileout, pAVIStreamOut2, strhdr)
    If (res <> AVIERR_OK) Then GoTo ErrorOut
    MsgBox "creado stream en el fichero (con el stream info del bitmap"
    
    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    Dim pOpts As Long
    Dim opts As AVI_COMPRESS_OPTIONS
    pOpts = VarPtr(opts)
    res = AVISaveOptions(Me.hWnd, _
                        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
                        1, _
                        pAVIStreamOut2, _
                        pOpts) 'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then 'In C TRUE = 1
        Call AVISaveOptionsFree(1, pOpts)
        GoTo ErrorOut
    End If
    MsgBox "relleno compress_options"
    'make compressed stream
    res = AVIMakeCompressedStream(pAVIStreamCompressed, pAVIStreamOut2, opts, 0&)
    If res <> AVIERR_OK Then GoTo ErrorOut
    MsgBox "creado stream comprimido desde stream de fichero con compress_options"
    
    'set format of stream according to the bitmap
'    With bih
'        .biBitCount = dib.BitCount
'        .biClrImportant = dib.ClrImportant
'        .biClrUsed = dib.ClrUsed
'        .biCompression = dib.Compression
'        .biHeight = dib.Height
'        .biWidth = dib.Width
'        .biPlanes = dib.Planes
'        .biSize = dib.SizeInfoHeader
'        .biSizeImage = dib.SizeImage
'        .biXPelsPerMeter = dib.XPPM
'        .biYPelsPerMeter = dib.YPPM
'    End With
'    MsgBox "relleno bitmap info header para comunicar con el stream"
    'set the format of the compressed stream
    res = AVIStreamSetFormat(pAVIStreamCompressed, 0, ByVal dib.PointerToBitmapInfo, dib.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo ErrorOut
    MsgBox "stream comprimido -> setformat"
    MsgBox "todo listo para escribir"
        
    'create a DIB class to load the frames into
    Set dib = New cDIB
    Dim j As Long 'contador de escritura
    j = 0
    
    Dim repetirvideo As Long
    repetirvideo = 2
    Do While (repetirvideo > 0)
        For i = firstFrame To (numFrames - 1) + firstFrame
            pDIB = AVIStreamGetFrame(pGetFrameObj, i) 'returns "packed DIB"
            If dib.CreateFromPackedDIBPointer(pDIB) Then
                res = AVIStreamWrite(pAVIStreamCompressed, _
                                    j, _
                                    1, _
                                    dib.PointerToBits, _
                                    dib.SizeImage, _
                                    AVIIF_KEYFRAME, _
                                    ByVal 0&, _
                                    ByVal 0&)
                
                If res <> AVIERR_OK Then GoTo ErrorOut
                j = j + 1
            Else
                
            End If
        Next
        repetirvideo = repetirvideo - 1
        MsgBox "copiado una vez (" & repetirvideo & ")"
    Loop
    
    Set dib = Nothing
    MsgBox "leidos bitmaps y escritos en streamOut"
        
              
        
    '// FIN
    MsgBox "todo correcto"
ErrorOut:
    Set dib = Nothing
    
    If (pAVIStreamCompressed <> 0) Then
        Call AVIStreamClose(pAVIStreamCompressed)
    End If

    If (pAVIStreamOut2 <> 0) Then
        Call AVIStreamClose(pAVIStreamOut2)
    End If

    If (pfileout <> 0) Then
        Call AVIFileClose(pfileout)
    End If
    
    If (pOpts <> 0) Then
        Call AVISaveOptionsFree(1, pOpts) 'frees resources
    End If

    If pGetFrameObj <> 0 Then
        Call AVIStreamGetFrameClose(pGetFrameObj) '//deallocates the GetFrame resources
    End If
    
    If pAVIStream <> 0 Then
        Call AVIStreamRelease(pAVIStream) '//closes video stream
    End If

     If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile) '// closes the file
    End If
    
    Call AVIFileExit '// releases AVIFile library
    
    If (Err <> AVIERR_OK) Then
        MsgBox "-- algun error --"
    End If
End Sub
