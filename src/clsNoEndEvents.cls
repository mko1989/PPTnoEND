VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNoEndEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Receives PowerPoint application events for the NoEnd add-in.
' The guard stops itself (from its timer) once no slide show is running.
Public WithEvents App As PowerPoint.Application

Private Sub App_SlideShowBegin(ByVal Wn As SlideShowWindow)
    NoEnd_StartGuard
End Sub
