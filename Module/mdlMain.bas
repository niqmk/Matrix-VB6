Attribute VB_Name = "mdlMain"
Option Explicit

Public Sub Main()
    Set mdlGlobal.fso = New FileSystemObject

    mdlGlobal.strPath = App.Path
    
    If Not Right(mdlGlobal.strPath, 1) = "\" Then mdlGlobal.strPath = mdlGlobal.strPath & "\"
    
    frmSplash.Show
End Sub
