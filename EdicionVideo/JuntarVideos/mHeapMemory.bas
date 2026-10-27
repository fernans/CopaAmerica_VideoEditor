Attribute VB_Name = "mHeapMemory"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (Destination As Any, _
    Source As Any, ByVal Length As Long)
Public Declare Function GetProcessHeap Lib "kernel32" () As Long
Public Declare Function HeapAlloc Lib "kernel32" _
    (ByVal hHeap As Long, ByVal dwFlags As Long, _
     ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" _
    (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Sub CopyMemoryWrite Lib "kernel32" Alias _
    "RtlMoveMemory" (ByVal Destination As Long, _
    Source As Any, ByVal Length As Long)
Public Declare Sub CopyMemoryRead Lib "kernel32" Alias _
    "RtlMoveMemory" (Destination As Any, _
    ByVal Source As Long, ByVal Length As Long)

