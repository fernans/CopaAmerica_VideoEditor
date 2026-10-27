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