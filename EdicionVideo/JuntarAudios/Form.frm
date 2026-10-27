VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JuntarVideos"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3870
   Icon            =   "Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'DefLng A-Z
Dim pAVIFileAudio As Long  'pointer to AVI File (PAVIFILE handle)
Dim pAVIStreamAudio As Long  'pointer to AVI stream interface (PAVISTREAM handle)
Dim numFramesAudio As Long 'number of frames in video stream
Dim firstFrameAudio As Long 'position of the first video frame
Dim pWaveFormat As WAVEFORMATEX
Dim pFormat As Long
Dim cbFormat As Long

Dim pfileoutAudio As Long
Dim pAVIStreamOutAudio As Long


Dim res As Long
    
Private Sub Form_Load()
  
    Dim hHeap As Long
    hHeap = GetProcessHeap
    '//pFrameformat = HeapAlloc(hHeap, 0, frameformatsize)     // asignar memoria
    
    Call AVIFileInit  '// opens AVIFile library
    
    Dim szFileOut As String
    Dim szFile As New Collection

    szFile.Add App.Path & "/source/audio4.wav"
    szFile.Add App.Path & "/source/audio5.wav"
    szFile.Add App.Path & "/source/audio6.wav"
    szFileOut = App.Path & "/target/output2.avi"
        
    '// abrir archivo 1.. para copiar formato
    AbrirArchivo szFile.Item(1)
    If res <> AVIERR_OK Then GoTo errorout
    
    MsgBox "todo listo para leer"
    
        
        
    '// abrir archivo para guardar (con el formato del 1er audio
    GuardarArchivo szFileOut
    If res <> AVIERROR_OK Then GoTo errorout
    
    MsgBox "todo listo para escribir"
        
        
    'Cerrar archivo In para empezar con la lista
    CerrarArchivoIn
        
        
    'create a DIB class to load the frames into
'    Set dib = New cDIB
    Dim bufAudio As Long
    Dim cbBufAudio As Long
    Dim bWritten As Long
    Dim sWritten As Long
    Dim i As Long 'contador de lectura
    Dim j As Long 'contador de escritura
    j = 0
   
    For Each szFileIn In szFile
            AbrirArchivo (szFileIn)
            If res <> AVIERR_OK Then GoTo errorout
    
        For i = 0 To numFramesAudio - 1 Step pWaveFormat.nSamplesPerSec
    
            'res = AVIStreamRead(pAVIStream, 0, numFrames, 0&, 0, cbBuf, 0&)
            res = AVIStreamRead(pAVIStreamAudio, i, pWaveFormat.nSamplesPerSec, 0&, 0, cbBufAudio, 0&)
           ''MsgBox i & "=" & cbBuf
            If res <> AVIERR_OK Then GoTo errorout
            bufAudio = HeapAlloc(hHeap, 0, cbBufAudio)
            If (bufAudio <> 0) Then
               ''MsgBox "allocated"
                'res = AVIStreamRead(pAVIStream, 0, numFrames, buf, cbBuf, bWritten, sWritten)
                res = AVIStreamRead(pAVIStreamAudio, i, pWaveFormat.nSamplesPerSec, bufAudio, cbBufAudio, bWritten, sWritten)
                If res <> AVIERR_OK Then GoTo errorout
               ''MsgBox "readed " & bWritten & " - " & sWritten
                res = AVIStreamWrite(pAVIStreamOutAudio, j, sWritten, bufAudio, cbBufAudio, 0, ByVal 0&, ByVal 0&) 'AVIIF_KEYFRAME
                If res <> AVIERR_OK Then GoTo errorout
               ''MsgBox "writed"

                HeapFree hHeap, 0, bufAudio
                bufAudio = 0
               ''MsgBox "free"
            End If

            'j = j + numFrames
            j = j + sWritten
            
        Next i
            
        'cerrar archivo lectura, para ir con el siguiente
        CerrarArchivoIn
        MsgBox "copiado video (" & szFileIn & ")"
    Next
           
    '// FIN
    MsgBox "todo correcto"
errorout:
'    Set dib = Nothing
    If (buf <> 0) Then
        HeapFree hHeap, 0, bufAudio
        bufAudio = 0
    End If
    
' ahora pFormat solo es el VarPtr de pWaveFormat
'    If (pFormat <> 0) Then
'        HeapFree hHeap, 0, pFormat
'        pFormat = 0
'    End If
    
    'Cerrar archivo de salida
    CerraArchivoOut
   
    'Cerrar archivo y stream entrada
    CerrarArchivoIn
    
    Call AVIFileExit '// releases AVIFile library
    
    If (res <> AVIERR_OK) Then
        MsgBox "-- algun error --"
    End If
End Sub

Private Sub CerrarArchivoIn()
    If pAVIStreamAudio <> 0 Then
        Call AVIStreamRelease(pAVIStreamAudio) '//closes video stream
        pAVIStreamAudio = 0
    End If

     If pAVIFileAudio <> 0 Then
        Call AVIFileRelease(pAVIFileAudio) '// closes the file
        pAVIFileAudio = 0
    End If
End Sub

Private Sub CerraArchivoOut()
    If (pAVIStreamOutAudio <> 0) Then
        Call AVIStreamClose(pAVIStreamOutAudio)
    End If

    If (pfileoutAudio <> 0) Then
        Call AVIFileClose(pfileoutAudio)
    End If
End Sub

Private Sub AbrirArchivo(szFile As String)
    res = AVIFileOpen(pAVIFileAudio, szFile, OF_SHARE_DENY_WRITE, 0&)
    If res <> AVIERR_OK Then GoTo errorout
    MsgBox "fichero abierto"
    
    '// leer stream
    'Get the first available video stream (PAVISTREAM)
    res = AVIFileGetStream(pAVIFileAudio, pAVIStreamAudio, streamtypeAUDIO, 0)
    If res <> AVIERR_OK Then GoTo errorout
    MsgBox "stream abierto"
    
               
    '// informacion del stream
    'get the starting position of the stream (some streams may not start simultaneously)
    firstFrameAudio = AVIStreamStart(pAVIStreamAudio)
    If firstFrameAudio = -1 Then GoTo errorout 'this function returns -1 on error
    'get the length of video stream in frames
    numFramesAudio = AVIStreamLength(pAVIStreamAudio)
    If numFramesAudio = -1 Then GoTo errorout ' this function returns -1 on error
    MsgBox "first: " & firstFrameAudio & " - num: " & numFramesAudio
    
    res = AVIStreamReadFormat(pAVIStreamAudio, 0, 0&, cbFormat)
    If res <> AVIERR_OK Then GoTo errorout
    MsgBox "allocated waveformat " & cbFormat & " -- " & Len(pWaveFormat)
        pFormat = VarPtr(pWaveFormat)
        res = AVIStreamReadFormat(pAVIStreamAudio, 0, pFormat, cbFormat)
        If res <> AVIERR_OK Then GoTo errorout
        MsgBox "leido waveformat"
    
'    Debug.Print pWaveFormat.wFormatTag
'    Debug.Print pWaveFormat.nChannels
'    Debug.Print pWaveFormat.nSamplesPerSec
'    Debug.Print pWaveFormat.nAvgBytesPerSec
'    Debug.Print pWaveFormat.nBlockAlign
'    Debug.Print pWaveFormat.wBitsPerSample
'    Debug.Print pWaveFormat.cbSize
'    MsgBox "cbformat " & cbFormat
errorout:
End Sub

Private Sub GuardarArchivo(szFileOut)
    '// crear un nuevo stream
    '// abrir archivo de salida
'    Dim pAVIStreamCompressed As Long
    '     Open the file for writing
    res = AVIFileOpen(pfileoutAudio, szFileOut, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo errorout
    MsgBox "nuevo avi abierto"
    ' Fill in the header for the video stream
    Dim strhdrAudio As AVI_STREAM_INFO
    With strhdrAudio
'        .fccType = mmioStringToFOURCC("vids", 0&)          '// stream type video
        .fccType = mmioStringToFOURCC("auds", 0&)          '// stream type video
        .fccHandler = 0&                                   '// default AVI handler
        .dwScale = pWaveFormat.nBlockAlign 'streamInfo.dwScale
        .dwRate = pWaveFormat.nSamplesPerSec * pWaveFormat.nBlockAlign 'streamInfo.dwRate                                        '// fps
'        .dwSuggestedBufferSize = dib.SizeImage             '// size of one frame pixels
        .dwSuggestedBufferSize = pWaveFormat.nBlockAlign 'streamInfo.dwSuggestedBufferSize               '// size of one frame pixels
        .dwSampleSize = pWaveFormat.nBlockAlign 'streamInfo.dwSampleSize
        Call SetRect(.rcFrame, 0, 0, 0, 0) '// rectangle for stream
        .dwQuality = -1
    End With
    
    MsgBox "stream info relleno del audio"
    ' And create the stream
    res = AVIFileCreateStream(pfileoutAudio, pAVIStreamOutAudio, strhdrAudio)
    If (res <> AVIERR_OK) Then GoTo errorout
    MsgBox "creado stream en el fichero (con el stream info del audio"
    
    
    'set the format of the compressed stream
    'pFormat = VarPtr(pWaveFormat)
    res = AVIStreamSetFormat(pAVIStreamOutAudio, 0, pWaveFormat, cbFormat)
    If (res <> AVIERR_OK) Then GoTo errorout
    MsgBox "stream salida -> setformat"
errorout:
End Sub
