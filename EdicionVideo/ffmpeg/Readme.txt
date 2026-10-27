FFMPEG Scout 1.00 Beta by Bytescout Software (www.bytescout.com)
----------------------

DESCRIPTION
===========

FFMPEG Scout is a helper library for free FFMpeg.exe open-source video files converter. FFMpeg Scout helps to convert AVI, MPEG into FLV through FFMpeg converter.

INSTALLATION
============

Unpack all files into the same folder.

Download and place FFMpeg.exe into the same folder so FFMpegScout.dll and FFMpeg.exe will be in the same folder

To download FFMPEG.exe use this links:
http://www.videohelp.com/tools?tool=ffmpeg
http://ffdshow.faireal.net/mirror/ffmpeg/

or use Google: http://www.google.com/search?q=ffmpeg.exe+download

Run Install.bat to install library

HOW TO USE
===========

Visual Basic/VBScript example code:

  W = 640
  H = 480
  Set VideoConverter = CreateObject("FFMpegScout.Application")
  VideoConverter.InputFileName = "testvideo.avi"
  VideoConverter.OutputFileName = "testvideo.flv"
  VideoConverter.Execute
 
  Set VideoConverter = Nothing

