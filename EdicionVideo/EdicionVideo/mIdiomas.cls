VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIdioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type t_idioma
    id As String
    ruta As String
End Type

Private pIdioma As t_idioma

Public Property Get id() As String
    id = pIdioma.id
End Property

Public Property Get ruta() As String
    ruta = pIdioma.ruta
End Property

Public Property Let id(nid As String)
    pIdioma.id = nid
End Property

Public Property Let ruta(nruta As String)
    pIdioma.ruta = nruta
End Property

