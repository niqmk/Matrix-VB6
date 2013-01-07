VERSION 5.00
Begin VB.Form frmLU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LU"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9150
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLU 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Index           =   0
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox txtLU 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Index           =   1
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdExportLU 
      Caption         =   "Eksport"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblLU 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lblLU 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      Height          =   195
      Index           =   1
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   120
   End
End
Attribute VB_Name = "frmLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If frmMain.Parent Then
        SetInitialization
    Else
        Me.txtLU(0).Text = frmMain.MatrixL
        Me.txtLU(1).Text = frmMain.MatrixU
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Parent = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLU = Nothing
End Sub

Private Sub cmdExportLU_Click()
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
            
            Dim intMatrixMode As Integer
            
            dblSource = frmMain.SourceMatrix
            
            intMatrixMode = frmMain.MatrixMode
            
            mdlFile.WriteTextLU .FileName, Me.txtLU(0).Text, Me.txtLU(1).Text, dblSource, intMatrixMode
        End If
    End With
    
ErrHandler:
End Sub

Private Sub SetInitialization()
    Dim dblSource() As Double
    Dim dblResult() As Double
    Dim dblResult2() As Double
    
    Dim intMatrixMode As Integer
    
    dblSource = frmMain.SourceMatrix
    
    intMatrixMode = frmMain.MatrixMode
    
    mdlLU.CalculateLU dblSource, dblResult, dblResult2, intMatrixMode
            
    mdlProcedures.SetMatrixToTextBox Me.txtLU(1), dblResult, intMatrixMode
    mdlProcedures.SetMatrixToTextBox Me.txtLU(0), dblResult2, intMatrixMode
End Sub
