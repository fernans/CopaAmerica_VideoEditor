' ////////////////////////////////////////////////
' converting AVI video into FLV using FFMPEG Scout
' ////////////////////////////////////////////////

  W = 640
  H = 480
  Set VideoConverter = CreateObject("FFMpegScout.Application")
  VideoConverter.InputFileName = "testvideo.avi"
  VideoConverter.OutputFileName = "testvideo.flv"
  VideoConverter.Execute
 
  Set VideoConverter = Nothing

' ////////////////////////////////////////////////
' now converting FLV video into SWF using SWF Scout
' Requires SWF Scout installed (http://bytescout.com/files/SWFScout.exe)
' ////////////////////////////////////////////////

  W = 640
  H = 480
  Set Movie = CreateObject("SWFScout.FlashMovie")
  Movie.InitLibrary "demo", "demo"

' Movie creating and setting parameters

  Movie.BeginMovie 0,0,W,H,1,12,8
  Movie.Compressed = true
  Movie.SetBackgroundColor 255,255,255

  Movie.AddPreloader 0, 0,0,0,180, 0,255,0,30, "Loading video (%d% completed)...", -1

 FontBig = Movie.AddFont("Arial",40,true,false,false,false,DEFAULT_CHARSET)


 Text = Movie.AddText2 ("Video",0,0,0,255,FontBig, W / 2, 60,2)
 Movie.PlaceText Text,Movie.CurrentMaxDepth ' place text

  Video= Movie.AddVideoFromFileName("testvideo.flv")
  Movie.FramesPerSecond= Movie.VIDEO_FramesPerSecond
  Movie.PlaceVideo Video,Movie.CurrentMaxDepth
  Movie.PLACE_SetTranslate 190,150
  Movie.ShowFrame Movie.VIDEO_FrameCount

 Movie.EndMovie

 Movie.SaveToFile "TestVideo.swf"

