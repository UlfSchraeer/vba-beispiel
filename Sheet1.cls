VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Function NeuerKontakt(ByVal Vorname As String, ByVal Nachname As String) As kontakt
    Dim k As New kontakt
    k.Vorname = Vorname
    k.Nachname = Nachname
    Set NeuerKontakt = k
End Function


Sub TestKontakt()
    Dim i As Integer
    Dim kontakte As New Collection
    
    kontakte.Add NeuerKontakt("Peter", "MŸller")
    kontakte.Add NeuerKontakt("Eva", "Schneider")
    kontakte.Add NeuerKontakt("Fritz", "Wagner")
    
    For i = 1 To kontakte.Count
        Me.Range("A" & i).Value = kontakte(i).VollerName
    Next i
    
End Sub


