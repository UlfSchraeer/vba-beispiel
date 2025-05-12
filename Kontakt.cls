VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Kontakt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pVorname As String
Private pNachname As String

Public Property Get Vorname() As String
 Vorname = pVorname
End Property

Public Property Let Vorname(ByVal Wert As String)
 pVorname = Wert
End Property

Public Property Get Nachname() As String
 Nachname = pNachname
End Property

Public Property Let Nachname(ByVal Wert As String)
 pNachname = Wert
End Property

Public Property Get VollerName() As String
 VollerName = pVorname + " " + pNachname
End Property
