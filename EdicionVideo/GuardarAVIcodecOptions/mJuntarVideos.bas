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



Public Function GrabarStream(szFileOut As String) As Long
    Dim mmBits() As Byte
    Dim nn As Long
    '// abrir archivo de salida
    '     Open the file for writing
    MsgBox "abrir"
    res = AVIFileOpen(pAVIFileOut, szFileOut, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo errorout
    MsgBox "abrir archivo"
    ' Fill in the header for the video stream
    With streamInfoOut
        .fccType = mmioStringToFOURCC("vids", 0&)          '// stream type video
        .fccHandler = 0&                                   '// default AVI handler
        .dwScale = 1
        .dwRate = 25                                       '// fps
        .dwSuggestedBufferSize = CLng(512) '//dib.SizeImage             '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, 320, 240) '// rectangle for stream
    End With
    
    '// crear un nuevo stream
    ' And create the stream
    res = AVIFileCreateStream(pAVIFileOut, pAVIStreamOut, streamInfoOut)
    If (res <> AVIERR_OK) Then GoTo errorout
    MsgBox "create stream"
    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
    res = AVISaveOptions(fPrincipal.hWnd, _
                        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
                        1, _
                        pAVIStreamOut, _
                        pOpts) 'returns TRUE if User presses OK, FALSE if Cancel, or error code
    
'    res = AVISaveOptions(fPrincipal.hWnd, _
'                        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
'                        1, _
'                        pAVIStreamOut, _
'                        pOpts) 'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then 'In C TRUE = 1
        Call AVISaveOptionsFree(1, pOpts)
        GoTo errorout
    End If
    MsgBox "save options"

Open App.Path & "\codec.ini" For Binary Access Write As 1
    Put 1, , opts.fccType
    Put 1, , opts.fccHandler
    Put 1, , opts.dwKeyFrameEvery
    Put 1, , opts.dwQuality
    Put 1, , opts.dwBytesPerSecond
    Put 1, , opts.dwFlags
    Put 1, , opts.cbFormat
    
    If (opts.cbFormat > 0) Then
        ReDim mmBits(0 To opts.cbFormat - 1)
        Call CopyMemory(mmBits(0), opts.lpFormat, opts.cbFormat)
        For nn = 0 To opts.cbFormat - 1
            Put 1, , mmBits(nn)
        Next
    End If
    
    Put 1, , opts.cbParms
    If (opts.cbParms > 0) Then
        ReDim mmBits(0 To opts.cbParms - 1)
        Call CopyMemory(mmBits(0), opts.lpParms, opts.cbParms)
        For nn = 0 To opts.cbParms - 1
            Put 1, , mmBits(nn)
        Next
    End If
    
    Put 1, , opts.dwInterleaveEvery

Close 1
  
    
'    .fccType = 0
'    .fccHandler = 2021026148
'    .dwKeyFrameEvery = 0
'    .dwQuality = 0
'    .dwBytesPerSecond = 0
'    .dwFlags = 8
'    .lpFormat = 0
'    .cbFormat = 0
'    .lpParms = 97087648
'    .cbParms = 512
'    .dwInterleaveEvery = 0


    'make compressed stream
    res = AVIMakeCompressedStream(pAVIStreamCompressed, pAVIStreamOut, opts, 0&)
    If res <> AVIERR_OK Then GoTo errorout
    
'    'set the format of the compressed stream
'    res = AVIStreamSetFormat(pAVIStreamCompressed, 0, ByVal dib.PointerToBitmapInfo, dib.SizeBitmapInfo)
'    If (res <> AVIERR_OK) Then GoTo errorout
    res = AVIERR_OK
errorout:
    GrabarStream = res
End Function

Public Sub CerrarTodo()
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
    
    Call AVIFileExit '// releases AVIFile library
    
    If (res <> AVIERR_OK) Then
        MsgBox "-- algun error --"
    End If
    
End Sub

Public Sub CerrarStreamIn()
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
