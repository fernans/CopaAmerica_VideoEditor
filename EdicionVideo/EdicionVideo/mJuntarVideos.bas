Attribute VB_Name = "mJuntarVideos"
'Option Explicit
'DefLng A-Z

'//Definicion de Variables Globales
Dim res As Long 'recibe resultado de llamadas API

Dim pAVIFileIn As Long  'pointer to AVI File (PAVIFILE handle)
Dim pAVIStreamIn As Long  'pointer to AVI stream interface (PAVISTREAM handle)
Dim fileInfo As AVI_FILE_INFO  'file info struct
Dim streamInfo As AVI_STREAM_INFO 'stream info struct
Dim numFrames As Long 'number of frames in video stream
Dim firstFrame As Long 'position of the first video frame

Dim dib As cDIB 'Buffer lector del Stream
Dim pGetFrameObj As Long 'pointer to GetFrame interface
Dim pDIB As Long 'pointer to packed DIB in memory
Dim bih As BITMAPINFOHEADER 'infoheader to pass to GetFrame functions

Dim pAVIFileOut As Long

Dim pAVIStreamOut As Long
Dim pAVIStreamCompressed As Long
Dim streamInfoOut As AVI_STREAM_INFO

Dim pOpts As Long
Dim opts As AVI_COMPRESS_OPTIONS

'//AUDIO
Dim pAVIFileAudio As Long  'pointer to AVI File (PAVIFILE handle)
Dim pAVIStreamAudio As Long  'pointer to AVI stream interface (PAVISTREAM handle)
Dim pAVIStreamOutAudio As Long
Dim numFramesAudio As Long 'number of frames in video stream
Dim firstFrameAudio As Long 'position of the first video frame
Dim pWaveFormat As WAVEFORMATEX
Dim pFormat As Long
Dim cbFormat As Long

Public Function JuntarVideos(szFileIn As Collection, szFileInAudio As Collection, szFileOut As String) As Long
    Dim ultimoEstado As String
    ultimoEstado = "sin empezar"

    '//Iniciar libreria
    Dim hHeap As Long
    hHeap = GetProcessHeap
    Call AVIFileInit  '// opens AVIFile library
    ultimoEstado = "call avifileinit"
    
    If (szFileOut = "") Then
        res = -1
        GoTo errorout
    End If
    If (szFileIn.Count < 1) Then
        res = -1
        GoTo errorout
    End If
    ultimoEstado = "comprobados filein fileout"
    
    Dim hayaudio As Boolean
    If (szFileInAudio.Count < 1) Then
        '// no hay audio
        hayaudio = False
    Else
        hayaudio = True
    End If
    ultimoEstado = "comprobado hay audio"
    
 '// VIDEO - leer formato //
    'abrir primer stream (se usará como formato para la salida)
    res = LeerStream(szFileIn.Item(1)) '// Devuelve el GetFrameObj para leer del stream
    If res <> AVIERR_OK Then GoTo errorout
    ultimoEstado = "leido primer filein"
    
    'leer el primer frame (se usará como formato para la salida)
    res = LeerDIB(firstFrame) '// Devuelve el dib relleno con el frame
    If res <> AVIERR_OK Then GoTo errorout
    ultimoEstado = "leido primer frame"
    
    'crear el archivo de salida para escribir
    res = GrabarStream(szFileOut) '// Devuelve el compressedStream para guardar en el
    If res <> AVIERR_OK Then GoTo errorout
    ultimoEstado = "creado archivo de salida"
    
    'cerrar stream inicial //preparando para inicio de bucle
    CerrarStreamIn
    ultimoEstado = "cerrado primer filein - info"
    
If (hayaudio) Then
 '// AUDIO - leer formato //
     '// abrir archivo 1.. para copiar formato
    AbrirArchivoAudio szFileInAudio.Item(1)
    If res <> AVIERR_OK Then GoTo errorout
    ultimoEstado = "abierto audio"
    
    '// abrir archivo para guardar (con el formato del 1er audio
    GuardarArchivoAudio pAVIFileOut
    If res <> AVIERROR_OK Then GoTo errorout
    ultimoEstado = "creado salida audio"
    
    'Cerrar archivo In para empezar con la lista
    CerrarArchivoInAudio
    ultimoEstado = "cerrado audio info"
End If 'hayaudio
 
 
    '// BUCLE PRINCIPAL //
    ' VIDEO:
    ' para cada nombre de archivo en lista de entrada
    ' abrir el video y obtener su stream
    ' recorrer los fotogramas del StreamIn hasta que se agote
    ' guardar el fotograma del StreamIn en el StreamOut (avanzar indice out)
    ' siguiente archivo (cerrar anterior, abrir siguiente, reiniciar indice StreamIn)
    ' cerrar streamIn
    ' cerrar streamOut
    ' AUDIO:
    ' para cada nombre de archivo en lista de entrada audio
    ' abrir el audio y obtener su stream
    ' recorrer por bloques de 1 seg. el stream hasta que se agote
    ' guardar el buffer del streamIn al streamOut (avanzar indice out)
    ' siguiente archivo audio (cerrar anterior, abrir siguiente, reiniciar indice StreamIn)
    ' FIN:
    ' vaciar objetos AVIFile
    ' cerrar libreria AVIFile

 '// VIDEO
    Dim archivoIn As String
    Dim indiceOut As Long
    indiceOut = 0
    ultimoEstado = "antes de comenzar bucle video"
    For i = 1 To szFileIn.Count
        archivoIn = szFileIn.Item(i)
        res = LeerStream(archivoIn) '// Devuelve el GetFrameObj para leer del stream
        ultimoEstado = "leido filein:" & i
        If res <> AVIERR_OK Then GoTo errorout
        'create a DIB class to load the frames into
        Set dib = New cDIB
        For j = firstFrame To (numFrames - 1) + firstFrame
            pDIB = AVIStreamGetFrame(pGetFrameObj, j) 'returns "packed DIB"
            If pDIB <> 0 Then
                If dib.CreateFromPackedDIBPointer(pDIB) Then
                    res = AVIStreamWrite(pAVIStreamCompressed, _
                                        indiceOut, _
                                        1, _
                                        dib.PointerToBits, _
                                        dib.SizeImage, _
                                        AVIIF_KEYFRAME, _
                                        ByVal 0&, _
                                        ByVal 0&)
                    If res <> AVIERR_OK Then GoTo errorout
                    indiceOut = indiceOut + 1
                Else
                  ' nada... me salto el frame
                End If
            End If
            fPrincipal.InformarCrearVideo CLng(((j * 100 / numFrames) / szFileIn.Count) + ((i - 1) * 100 / szFileIn.Count))
            fPrincipal.Refresh
            DoEvents
        Next    '// bucle fotogramas
        ultimoEstado = "fin bucle fotogramas"
        'cerrar streamIn //preparar para siguiente o fin
        CerrarStreamIn
        ultimoEstado = "cerrado filein:" & i
    Next '// bucle archivos
    Set dib = Nothing
    ultimoEstado = "fin bucle filein"
                  
If (hayaudio) Then
 '//AUDIO
    Dim bufAudio As Long
    Dim cbBufAudio As Long
    Dim bWritten As Long
    Dim sWritten As Long
    indiceOut = 0
    ultimoEstado = "antes bucle audio"
    
    For i = 1 To szFileInAudio.Count
        archivoIn = szFileInAudio.Item(i)
        AbrirArchivoAudio archivoIn
        If res <> AVIERR_OK Then GoTo errorout
        ultimoEstado = "abierto audio:" & i & " __ " & archivoIn
        For j = 0 To numFramesAudio - 1 Step pWaveFormat.nSamplesPerSec
            res = AVIStreamRead(pAVIStreamAudio, j, pWaveFormat.nSamplesPerSec, 0&, 0, cbBufAudio, 0&)
            If res <> AVIERR_OK Then GoTo errorout
            bufAudio = HeapAlloc(hHeap, 0, cbBufAudio)
            If (bufAudio <> 0) Then
                res = AVIStreamRead(pAVIStreamAudio, j, pWaveFormat.nSamplesPerSec, bufAudio, cbBufAudio, bWritten, sWritten)
                If res <> AVIERR_OK Then GoTo errorout
                res = AVIStreamWrite(pAVIStreamOutAudio, indiceOut, sWritten, bufAudio, cbBufAudio, 0, ByVal 0&, ByVal 0&) 'AVIIF_KEYFRAME
                If res <> AVIERR_OK Then GoTo errorout

                HeapFree hHeap, 0, bufAudio
                bufAudio = 0
                
                indiceOut = indiceOut + sWritten
            End If
            fPrincipal.InformarCrearVideo CLng(((j * 100 / numFramesAudio) / szFileInAudio.Count) + ((i - 1) * 100 / szFileInAudio.Count))
            fPrincipal.Refresh
            DoEvents
        Next    '//bucle de samples
        ultimoEstado = "fin bucle samples"
        'cerrar archivo lectura, para ir con el siguiente
        CerrarArchivoInAudio
    Next '// bucle de archivos
    ultimoEstado = "fin bucle audio"
 '// FIN AUDIO
End If 'hayaudio
 
 
    '// FIN
    res = AVIERR_OK
    ultimoEstado = "todo OK"
errorout:
    CerrarTodo
    If (res <> AVIERR_OK) Then MsgBox ultimoEstado
    JuntarVideos = res
End Function

Private Function LeerStream(szFile As String) As Long
    '// abrir archivo
    res = AVIFileOpen(pAVIFileIn, szFile, OF_SHARE_DENY_WRITE, 0&)
    If res <> AVIERR_OK Then GoTo errorout
    
    '// leer stream
    'Get the first available video stream (PAVISTREAM)
    res = AVIFileGetStream(pAVIFileIn, pAVIStreamIn, streamtypeVIDEO, 0)
    If res <> AVIERR_OK Then GoTo errorout
        
    '// avi info
    'get file info struct (UDT)
    res = AVIFileInfo(pAVIFileIn, fileInfo, Len(fileInfo))
    If res <> AVIERR_OK Then GoTo errorout

    '// stream info
    'get stream info struct (UDT)
    res = AVIStreamInfo(pAVIStreamIn, streamInfo, Len(streamInfo))
    If res <> AVIERR_OK Then GoTo errorout
        
    '// informacion del stream
    'get the starting position of the stream (some streams may not start simultaneously)
    firstFrame = AVIStreamStart(pAVIStreamIn)
    If firstFrame = -1 Then GoTo errorout 'this function returns -1 on error
    'get the length of video stream in frames
    numFrames = AVIStreamLength(pAVIStreamIn)
    If numFrames = -1 Then GoTo errorout ' this function returns -1 on error
          
    '// abrir frame
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
    Dim i As Long
    'init AVISTreamGetFrame* functions and create GETFRAME object
    pGetFrameObj = AVIStreamGetFrameOpen(pAVIStreamIn, bih) 'force function to return 24bit DIBS
    
    res = AVIERR_OK
errorout:
    LeerStream = res
End Function

Private Function LeerDIB(i As Long) As Long
    'create a DIB class to load the frames into
    Set dib = New cDIB
        pDIB = AVIStreamGetFrame(pGetFrameObj, i) 'returns "packed DIB"
        If dib.CreateFromPackedDIBPointer(pDIB) Then
            res = AVIERR_OK
        Else
            res = -1
        End If
    LeerDIB = res
End Function



Private Function GrabarStream(szFileOut As String) As Long
    Dim mmBitsFormat() As Byte
    Dim mmBitsParms() As Byte
    Dim nn As Long
    '// abrir archivo de salida
    '     Open the file for writing
    res = AVIFileOpen(pAVIFileOut, szFileOut, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo errorout
    ' Fill in the header for the video stream
    With streamInfoOut
        .fccType = mmioStringToFOURCC("vids", 0&)          '// stream type video
        .fccHandler = 0&                                   '// default AVI handler
        .dwScale = 1
        .dwRate = 25                                       '// fps
        .dwSuggestedBufferSize = dib.SizeImage             '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, dib.Width, dib.Height) '// rectangle for stream
    End With
    
    '// crear un nuevo stream
    ' And create the stream
    res = AVIFileCreateStream(pAVIFileOut, pAVIStreamOut, streamInfoOut)
    If (res <> AVIERR_OK) Then GoTo errorout
    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
'    pOpts = VarPtr(opts)
'    res = AVISaveOptions(Form1.hWnd, _
'                        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
'                        1, _
'                        pAVIStreamOut, _
'                        pOpts) 'returns TRUE if User presses OK, FALSE if Cancel, or error code
'    If res <> 1 Then 'In C TRUE = 1
'        Call AVISaveOptionsFree(1, pOpts)
'        GoTo errorout
'    End If

Open App.Path & "\codec.ini" For Binary Access Read As 1
    Get 1, , opts.fccType
    Get 1, , opts.fccHandler
    Get 1, , opts.dwKeyFrameEvery
    Get 1, , opts.dwQuality
    Get 1, , opts.dwBytesPerSecond
    Get 1, , opts.dwFlags
    
    Get 1, , opts.cbFormat
    opts.lpFormat = 0
    If (opts.cbFormat > 0) Then
        ReDim mmBitsFormat(0 To opts.cbFormat - 1)
        For nn = 0 To opts.cbFormat - 1
            Get 1, , mmBitsFormat(nn)
        Next
        opts.lpFormat = VarPtr(mmBitsFormat(0))
        'Call CopyMemory(mmBits(0), opts.lpFormat, opts.cbFormat)
    End If
    
    Get 1, , opts.cbParms
    opts.lpParms = 0
    If (opts.cbParms > 0) Then
        ReDim mmBitsParms(0 To opts.cbParms - 1)
        For nn = 0 To opts.cbParms - 1
            Get 1, , mmBitsParms(nn)
        Next
        opts.lpParms = VarPtr(mmBitsParms(0))
        'Call CopyMemory(mmBits(0), opts.lpParms, opts.cbParms)
    End If
    
    Get 1, , opts.dwInterleaveEvery

Close 1

    'make compressed stream
    res = AVIMakeCompressedStream(pAVIStreamCompressed, pAVIStreamOut, opts, 0&)
    If res <> AVIERR_OK Then GoTo errorout
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(pAVIStreamCompressed, 0, ByVal dib.PointerToBitmapInfo, dib.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo errorout
    res = AVIERR_OK
errorout:
    GrabarStream = res
End Function

Private Sub CerrarTodo()
    Set dib = Nothing
    If (pAVIStreamCompressed <> 0) Then
        Call AVIStreamClose(pAVIStreamCompressed)
        pAVIStreamCompressed = 0
    End If
    If (pAVIStreamOut <> 0) Then
        Call AVIStreamClose(pAVIStreamOut)
        pAVIStreamOut = 0
    End If
    If (pAVIFileOut <> 0) Then
        Call AVIFileClose(pAVIFileOut)
        pAVIFileOut = 0
    End If
    If (pOpts <> 0) Then
        Call AVISaveOptionsFree(1, pOpts) 'frees resources
        pOtps = 0
    End If

    CerrarStreamIn
    
    '//AUDIO
    CerraArchivoOutAudio
    CerrarArchivoInAudio
    
    Call AVIFileExit '// releases AVIFile library
    
'    If (res <> AVIERR_OK) Then
'        MsgBox "-- algun error --"
'    End If
    
End Sub

Private Sub CerrarStreamIn()
    If pGetFrameObj <> 0 Then
        Call AVIStreamGetFrameClose(pGetFrameObj) '//deallocates the GetFrame resources
        pGetFrameObj = 0
    End If
    If pAVIStreamIn <> 0 Then
        Call AVIStreamRelease(pAVIStreamIn) '//closes video stream
        pAVIStreamIn = 0
    End If
    If pAVIFileIn <> 0 Then
        Call AVIFileRelease(pAVIFileIn) '// closes the file
        pAVIFileIn = 0
    End If
End Sub


Private Sub CerrarArchivoInAudio()
    If pAVIStreamAudio <> 0 Then
        Call AVIStreamRelease(pAVIStreamAudio) '//closes video stream
        pAVIStreamAudio = 0
    End If

     If pAVIFileAudio <> 0 Then
        Call AVIFileRelease(pAVIFileAudio) '// closes the file
        pAVIFileAudio = 0
    End If
End Sub

Private Sub CerraArchivoOutAudio()
    '//cierro el stream... el avi se cierra con el video
    If (pAVIStreamOutAudio <> 0) Then
        Call AVIStreamClose(pAVIStreamOutAudio)
    End If
End Sub

Private Sub AbrirArchivoAudio(szFile As String)
    res = AVIFileOpen(pAVIFileAudio, szFile, OF_SHARE_DENY_WRITE, 0&)
    If res <> AVIERR_OK Then GoTo errorout
    
    '// leer stream
    'Get the first available video stream (PAVISTREAM)
    res = AVIFileGetStream(pAVIFileAudio, pAVIStreamAudio, streamtypeAUDIO, 0)
    If res <> AVIERR_OK Then GoTo errorout
               
    '// informacion del stream
    'get the starting position of the stream (some streams may not start simultaneously)
    firstFrameAudio = AVIStreamStart(pAVIStreamAudio)
    If firstFrameAudio = -1 Then GoTo errorout 'this function returns -1 on error
    'get the length of video stream in frames
    numFramesAudio = AVIStreamLength(pAVIStreamAudio)
    If numFramesAudio = -1 Then GoTo errorout ' this function returns -1 on error
    
    res = AVIStreamReadFormat(pAVIStreamAudio, 0, 0&, cbFormat)
    If res <> AVIERR_OK Then GoTo errorout
        pFormat = VarPtr(pWaveFormat)
        res = AVIStreamReadFormat(pAVIStreamAudio, 0, pFormat, cbFormat)
        If res <> AVIERR_OK Then GoTo errorout

errorout:
End Sub

Private Sub GuardarArchivoAudio(pfileoutAudio As Long)
    '// archivo de salida abierto en pfileoutAudio
    ' Fill in the header for the video stream
    Dim strhdrAudio As AVI_STREAM_INFO
    With strhdrAudio
        .fccType = mmioStringToFOURCC("auds", 0&)          '// stream type video
        .fccHandler = 0&                                   '// default AVI handler
        .dwScale = pWaveFormat.nBlockAlign                 '// un bloque
        .dwRate = pWaveFormat.nSamplesPerSec * pWaveFormat.nBlockAlign '//samples por segundo * tamaño del bloque
        .dwSuggestedBufferSize = pWaveFormat.nBlockAlign   '// un bloque
        .dwSampleSize = pWaveFormat.nBlockAlign            '// un bloque
        Call SetRect(.rcFrame, 0, 0, 0, 0) '// rectangle for stream
        .dwQuality = -1
    End With
    
    '// crear un nuevo stream
    res = AVIFileCreateStream(pfileoutAudio, pAVIStreamOutAudio, strhdrAudio)
    If (res <> AVIERR_OK) Then GoTo errorout
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(pAVIStreamOutAudio, 0, pWaveFormat, cbFormat)
    If (res <> AVIERR_OK) Then GoTo errorout

errorout:
End Sub

