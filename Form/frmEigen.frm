VERSION 5.00
Begin VB.Form frmEigen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EIGENVALUES dan EIGENVECTORS"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5655
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExportEigen 
      Caption         =   "Eksport"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox txtEigen 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Index           =   0
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.TextBox txtEigen 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Index           =   1
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label lblEigen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EigenValue"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblEigen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EigenVector"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   870
   End
End
Attribute VB_Name = "frmEigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If frmMain.Parent Then
        SetInitialization
    Else
        Me.txtEigen(0).Text = frmMain.EigenValues
        Me.txtEigen(1).Text = frmMain.EigenVectors
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Parent = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEigen = Nothing
End Sub

Private Sub cmdExportEigen_Click()
    On Local Error GoTo ErrHandler
    
    With frmMain.cdlFile
        .InitDir = mdlGlobal.strPath
        .Filter = "Text Files (*.txt)|*.txt"
        
        .ShowSave
        
        If Not Trim(.FileName) = "" Then
            If mdlGlobal.fso.FileExists(.FileName) Then
                mdlGlobal.fso.DeleteFile .FileName, True
            End If
            
            Dim dblSource() As Double
            
            dblSource = frmMain.SourceMatrix
            
            mdlFile.WriteTextEigen .FileName, Me.txtEigen(0).Text, Me.txtEigen(1).Text, dblSource, frmMain.MatrixMode
        End If
    End With
    
ErrHandler:
End Sub

Private Sub SetInitialization()
    Dim dblSource() As Double
    Dim dblResult() As Double
    
    dblSource = frmMain.SourceMatrix
    
    mdlEigen.CalculateEigen dblSource, frmMain.MatrixMode
    
    Dim dblEigenValues() As Double
    Dim dblEigenVectors() As Double
    
    dblEigenValues = mdlEigen.EigenValues
    dblEigenVectors = mdlEigen.EigenVectors
    
    Me.txtEigen(0).Text = ""
    Me.txtEigen(1).Text = ""
    
    Dim intCounter As Integer

    For intCounter = 0 To frmMain.MatrixMode - 1
        If Not intCounter = 0 Then
            Me.txtEigen(0).Text = Me.txtEigen(0).Text & vbCrLf
        End If

        Me.txtEigen(0).Text = Me.txtEigen(0).Text & "LAMBDA " & intCounter + 1 & " : " & dblEigenValues(intCounter)
    Next intCounter
    
    mdlProcedures.SetMatrixToTextBox Me.txtEigen(1), dblEigenVectors, frmMain.MatrixMode
End Sub
