VERSION 5.00
Begin VB.Form frmSVD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Singular Value Decomposition"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12120
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSVD 
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
      Index           =   2
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtSVD 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Index           =   1
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdExportSVD 
      Caption         =   "Eksport"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtSVD 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Index           =   0
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblSVD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vt"
      Height          =   195
      Index           =   2
      Left            =   8040
      TabIndex        =   6
      Top             =   120
      Width           =   150
   End
   Begin VB.Label lblSVD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   195
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   105
   End
   Begin VB.Label lblSVD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   105
   End
End
Attribute VB_Name = "frmSVD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If frmMain.Parent Then
        SetInitialization
    Else
        Me.txtSVD(0).Text = frmMain.MatrixA
        Me.txtSVD(1).Text = frmMain.MatrixE
        Me.txtSVD(2).Text = frmMain.MatrixVT
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Parent = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSVD = Nothing
End Sub

Private Sub cmdExportSVD_Click()
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
            
            mdlFile.WriteTextSVD .FileName, Me.txtSVD(0).Text, Me.txtSVD(1).Text, Me.txtSVD(2).Text, dblSource, intMatrixMode
        End If
    End With
    
ErrHandler:
End Sub

Private Sub SetInitialization()
    Dim dblSource() As Double
    
    Dim intMatrixMode As Integer
    
    dblSource = frmMain.SourceMatrix
    
    intMatrixMode = frmMain.MatrixMode
    
    mdlSVD.CalculateSVD dblSource, intMatrixMode
    
    Dim dblA() As Double
    Dim dblE() As Double
    Dim dblVT() As Double
    
    dblA = mdlSVD.MatrixA
    dblE = mdlSVD.MatrixE
    dblVT = mdlProcedures.SetTranspose2(mdlSVD.MatrixVT)
    
    Me.txtSVD(0).Text = ""
    Me.txtSVD(1).Text = ""
    Me.txtSVD(2).Text = ""
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To UBound(dblA, 2)
        For intColumn = 0 To UBound(dblA, 1)
            Me.txtSVD(0).Text = Me.txtSVD(0).Text & Format(dblA(intColumn, intRow), "0.00") & vbTab
        Next intColumn
        
        Me.txtSVD(0).Text = Left(Me.txtSVD(0).Text, Len(Me.txtSVD(0).Text) - 1)
        
        Me.txtSVD(0).Text = Me.txtSVD(0).Text & vbCrLf
    Next intRow
    
    If Len(Me.txtSVD(0).Text) > 0 Then Me.txtSVD(0).Text = Left(Me.txtSVD(0).Text, Len(Me.txtSVD(0).Text) - 2)
    
    mdlProcedures.SetMatrixToTextBox Me.txtSVD(1), dblE, intMatrixMode
    
    For intRow = 0 To UBound(dblVT, 1)
        For intColumn = 0 To UBound(dblVT, 2)
            Me.txtSVD(2).Text = Me.txtSVD(2).Text & Format(dblVT(intRow, intColumn), "0.00") & vbTab
        Next intColumn
        
        Me.txtSVD(2).Text = Left(Me.txtSVD(2).Text, Len(Me.txtSVD(2).Text) - 1)
        
        Me.txtSVD(2).Text = Me.txtSVD(2).Text & vbCrLf
    Next intRow
    
    If Len(Me.txtSVD(2).Text) > 0 Then Me.txtSVD(2).Text = Left(Me.txtSVD(2).Text, Len(Me.txtSVD(2).Text) - 2)
End Sub
