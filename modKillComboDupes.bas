Attribute VB_Name = "modKillComboDupes"
' Name  : Remove Duplicate ComboBox Items
' By    : Rudy Alex Kohn [rudyalexkohn@hotmail.com]
Option Explicit

Sub KillCmbDupes(cmb As ComboBox, Optional DoSecondPass As Boolean = True)
' This should (hopefully) remove any dupes that exists in a combo..

  On Error Resume Next
  Dim i As Integer
  Dim x As Integer
  For i = 0 To cmb.ListCount - 1  ' Pass 1
    For x = 0 To cmb.ListCount - 1
      If i <> x Then If LCase$(cmb.List(x)) = LCase$(cmb.List(i)) Then cmb.RemoveItem x
    Next
  Next

  If (DoSecondPass) Then
    ' 2 Passes needed in some cases
    For i = 0 To cmb.ListCount - 1  ' Pass 2
      For x = 0 To cmb.ListCount - 1
        If i <> x Then If LCase$(cmb.List(x)) = LCase$(cmb.List(i)) Then cmb.RemoveItem x
      Next
    Next
  End If
End Sub

