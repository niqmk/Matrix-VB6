VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TOOLBOX MATRIX"
   ClientHeight    =   7950
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMatrix 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.Frame fraResult 
      Height          =   2415
      Left            =   1080
      TabIndex        =   6
      Top             =   5280
      Width           =   10215
      Begin VB.TextBox txtNote 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   8895
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame fraMatrix 
      Height          =   1215
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   10215
      Begin MSComDlg.CommonDialog cdlFile 
         Left            =   3480
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.ComboBox cmbMatrix 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblMatrix 
         Caption         =   "Masukkan Ukuran Matrix"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxMatrix 
      Height          =   3375
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5953
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   3  'Align Left
      Height          =   7950
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   14023
      ButtonWidth     =   2831
      ButtonHeight    =   1429
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "imlMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList imlMain 
         Left            =   120
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":68A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":78F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":9D48
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":19D9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C1EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1E63E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":20A90
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxInverse 
      Height          =   3375
      Left            =   6240
      TabIndex        =   9
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5953
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Matrix Inverse"
      Height          =   195
      Index           =   1
      Left            =   6240
      TabIndex        =   11
      Top             =   1560
      Width           =   990
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Matrix Input"
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   1560
      Width           =   825
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
         Begin VB.Menu mnuLU 
            Caption         =   "LU"
         End
         Begin VB.Menu mnuEigen 
            Caption         =   "Eigen"
         End
         Begin VB.Menu mnuSVD 
            Caption         =   "SVD"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbox 
         Caption         =   "&Toolbox"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ToolbarConst
    MatrixTl = 1
    OpenTl
    ClearTl
    CheckTl
    PivotTl
    EigenTl
    LUTl
    SVDTl
    CloseTl
End Enum

Private objToolbarConst As ToolbarConst

Private strL As String
Private strU As String
Private strEigenValues As String
Private strEigenVectors As String
Private strA As String
Private strE As String
Private strVT As String

Private blnMatrix As Boolean
Private blnFocus As Boolean

Private blnNonSingular As Boolean
Private blnSquare As Boolean
Private blnSymmetric As Boolean

Private blnParent As Boolean

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mdlGlobal.fso = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuHelp_Click()
    If mdlGlobal.fso.FileExists(mdlGlobal.strPath & "help.chm") Then
        mdlProcedures.ShellExecute Me.hwnd, "", mdlGlobal.strPath & "help.chm", "", "", 1
    End If
End Sub

Private Sub mnuNew_Click()
    NewFile
End Sub

Private Sub mnuOpen_Click()
    OpenFile
End Sub

Private Sub mnuSave_Click()
    SaveFile
End Sub

Private Sub mnuPrint_Click()
    On Local Error GoTo ErrHandler
    
    If Not blnMatrix Then Exit Sub
    
    Dim dblSource() As Double
    Dim dblInverse() As Double
    
    dblSource = Me.SourceMatrix
    dblInverse = Me.InverseMatrix
    
    mdlExcel.PrintMatrix dblSource, dblInverse, Me.cmbMatrix.ListIndex + 2
    
    If mdlGlobal.fso.FileExists(mdlGlobal.strPath & "Print.xls") Then
        mdlGlobal.fso.DeleteFile mdlGlobal.strPath & "Print.xls"
    End If
    
ErrHandler:
End Sub

Private Sub mnuLU_Click()
    On Local Error GoTo ErrHandler
    
    With Me.cdlFile
        .InitDir = mdlGlobal.strPath
        .Filter = "Text Files (*.txt)|*.txt"
        
        .ShowOpen
        
        If Not Trim(.FileName) = "" Then
            If Not mdlGlobal.fso.FileExists(.FileName) Then
                MsgBox "Data Tidak Ada", vbOKOnly + vbExclamation, Me.Caption
                
                Exit Sub
            End If
            
            Dim dblSource() As Double
            
            Dim intMatrixMode As Integer
            
            intMatrixMode = mdlFile.ReadTextLU(.FileName, strL, strU, dblSource)
            
            If Not intMatrixMode = 0 Then
                frmMain.MatrixMode = intMatrixMode
                
                mdlProcedures.SetMatrixToFlexGrid frmMain.flxMatrix, dblSource, intMatrixMode
                
                objToolbarConst = CheckTl
                
                ShowMatrix True
                
                Dim dblResult() As Double
                
                dblResult = mdlProcedures.SetInverse(dblSource, intMatrixMode)
                
                mdlProcedures.SetMatrixToFlexGrid frmMain.flxInverse, dblResult, intMatrixMode
                
                frmLU.Show vbModal
            End If
        End If
    End With
    
ErrHandler:
End Sub

Public Property Get MatrixU() As String
    MatrixU = strU
End Property

Public Property Get MatrixL() As String
    MatrixL = strL
End Property

Private Sub mnuEigen_Click()
    On Local Error GoTo ErrHandler
    
    With Me.cdlFile
        .InitDir = mdlGlobal.strPath
        .Filter = "Text Files (*.txt)|*.txt"
        
        .ShowOpen
        
        If Not Trim(.FileName) = "" Then
            If Not mdlGlobal.fso.FileExists(.FileName) Then
                MsgBox "Data Tidak Ada", vbOKOnly + vbExclamation, Me.Caption
                
                Exit Sub
            End If
            
            Dim dblSource() As Double
            
            Dim intMatrixMode As Integer
            
            intMatrixMode = mdlFile.ReadTextEigen(.FileName, strEigenValues, strEigenVectors, dblSource)
            
            If Not intMatrixMode = 0 Then
                Me.MatrixMode = intMatrixMode
                
                mdlProcedures.SetMatrixToFlexGrid Me.flxMatrix, dblSource, intMatrixMode
                
                objToolbarConst = CheckTl
                
                ShowMatrix True
                
                Dim dblResult() As Double
                
                dblResult = mdlProcedures.SetInverse(dblSource, intMatrixMode)
                
                mdlProcedures.SetMatrixToFlexGrid frmMain.flxInverse, dblResult, intMatrixMode
                
                frmEigen.Show vbModal
            End If
        End If
    End With
    
ErrHandler:
End Sub

Public Property Get EigenValues() As String
    EigenValues = strEigenValues
End Property

Public Property Get EigenVectors() As String
    EigenVectors = strEigenVectors
End Property

Private Sub mnuSVD_Click()
    On Local Error GoTo ErrHandler
    
    With frmMain.cdlFile
        .InitDir = mdlGlobal.strPath
        .Filter = "Text Files (*.txt)|*.txt"
        
        .ShowOpen
        
        If Not Trim(.FileName) = "" Then
            If Not mdlGlobal.fso.FileExists(.FileName) Then
                MsgBox "Data Tidak Ada", vbOKOnly + vbExclamation, Me.Caption
                
                Exit Sub
            End If
            
            Dim dblSource() As Double
            
            Dim intMatrixMode As Integer
            
            intMatrixMode = mdlFile.ReadTextSVD(.FileName, strA, strE, strVT, dblSource)
            
            If Not intMatrixMode = 0 Then
                frmMain.MatrixMode = intMatrixMode
                
                mdlProcedures.SetMatrixToFlexGrid frmMain.flxMatrix, dblSource, intMatrixMode
                
                objToolbarConst = CheckTl
                
                ShowMatrix True
                
                Dim dblResult() As Double
                
                dblResult = mdlProcedures.SetInverse(dblSource, intMatrixMode)
                
                mdlProcedures.SetMatrixToFlexGrid frmMain.flxInverse, dblResult, intMatrixMode
                
                frmSVD.Show vbModal
            End If
        End If
    End With
    
ErrHandler:
End Sub

Public Property Get MatrixA() As String
    MatrixA = strA
End Property

Public Property Get MatrixE() As String
    MatrixE = strE
End Property

Public Property Get MatrixVT() As String
    MatrixVT = strVT
End Property

Private Sub mnuToolbox_Click()
    Me.mnuToolbox.Checked = Not Me.mnuToolbox.Checked
    
    Me.tlbMain.Visible = Me.mnuToolbox.Checked
End Sub

Private Sub mnuStatusBar_Click()
    Me.mnuStatusBar.Checked = Not Me.mnuStatusBar.Checked
    
    Me.txtNote.Visible = Me.mnuStatusBar.Checked
    Me.lblNote.Visible = Me.mnuStatusBar.Checked
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    blnParent = True
    
    If Me.txtMatrix.Visible Then Me.txtMatrix.Visible = False
    
    Select Case Button.Index
        Case MatrixTl:
            blnMatrix = True
            
            ShowMatrix
        Case OpenTl
            OpenFile
        Case ClearTl
            NewFile
        Case CheckTl
            If blnMatrix Then
                objToolbarConst = CheckTl
                
                ShowMatrix True
            End If
        Case PivotTl
            If blnMatrix Then
                objToolbarConst = PivotTl
                
                ShowMatrix True
            End If
        Case EigenTl
            If blnMatrix Then
                objToolbarConst = EigenTl
                
                ShowMatrix True
            End If
        Case LUTl
            If blnMatrix Then
                objToolbarConst = LUTl
                
                ShowMatrix True
            End If
        Case SVDTl
            If blnMatrix Then
                objToolbarConst = SVDTl
                
                ShowMatrix True
            End If
        Case CloseTl
            Unload Me
    End Select
End Sub

Private Sub cmbMatrix_Click()
    ArrangeGrid Me.cmbMatrix.ListIndex + 2
    
    blnNonSingular = False
    blnSquare = False
    blnSymmetric = False
End Sub

Private Sub flxMatrix_DblClick()
    blnFocus = True
    
    ArrangeTextbox
End Sub

Private Sub flxMatrix_Scroll()
    Me.txtMatrix.Visible = False
End Sub

Private Sub flxMatrix_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        blnFocus = True
        
        ArrangeTextbox
    ElseIf (KeyCode >= vbKey0 And KeyCode <= vbKey9) Then
        blnFocus = False
        
        ArrangeTextbox
        
        Me.txtMatrix.Text = Chr(KeyCode)
    End If
End Sub

Private Sub txtMatrix_Change()
    Me.txtMatrix.SelStart = Len(Me.txtMatrix.Text)
End Sub

Private Sub txtMatrix_GotFocus()
    If blnFocus Then
        Me.txtMatrix.SelStart = 0
        Me.txtMatrix.SelLength = Len(Me.txtMatrix.Text)
    End If
End Sub

Private Sub txtMatrix_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        FillTextInput
        
        Me.txtMatrix.Visible = False
    ElseIf KeyCode = vbKeyEscape Then
        Me.txtMatrix.Visible = False
    End If
End Sub

Private Sub txtMatrix_LostFocus()
    Me.txtMatrix.Visible = False
End Sub

Private Sub SetInitialization()
    ArrangeGrid
    
    Dim intCounter As Integer
    
    With Me.cmbMatrix
        For intCounter = 2 To mdlGlobal.intMatrix
            .AddItem intCounter & " x " & intCounter
        Next intCounter
        
        .ListIndex = 0
    End With
    
    With Me.tlbMain
        .Buttons(MatrixTl).ToolTipText = "Buka Matrix"
        .Buttons(OpenTl).ToolTipText = "Buka File"
        .Buttons(ClearTl).ToolTipText = "Hapus"
        .Buttons(CheckTl).ToolTipText = "Check Matrix"
        .Buttons(PivotTl).ToolTipText = "Pivoting"
        .Buttons(EigenTl).ToolTipText = "Eigen"
        .Buttons(LUTl).ToolTipText = "LU"
        .Buttons(SVDTl).ToolTipText = "SVD"
        .Buttons(CloseTl).ToolTipText = "Keluar"
    End With
    
    Me.txtMatrix.Visible = False
    
    blnMatrix = False
    blnNonSingular = False
    blnSquare = False
    
    ShowMatrix
    
    ArrangeGrid Me.cmbMatrix.ListIndex + 2
End Sub

Private Sub ShowMatrix(Optional ByVal blnResult As Boolean = False)
    Me.fraMatrix.Visible = blnMatrix
    Me.flxMatrix.Visible = blnMatrix
    Me.flxInverse.Visible = blnMatrix
    Me.lblGrid(0).Visible = blnMatrix
    Me.lblGrid(1).Visible = blnMatrix
    
    Me.fraResult.Visible = blnResult
    
    If blnResult Then
        Dim dblSource() As Double
        Dim dblResult() As Double
        
        Dim intMatrixMode As Integer
        
        dblSource = Me.SourceMatrix
        
        intMatrixMode = Me.MatrixMode
        
        If objToolbarConst = CheckTl Then
            Dim dblDeterminant As Double
            
            dblDeterminant = mdlProcedures.GetDeterminant(dblSource, intMatrixMode)
            
            MsgBox "Nilai Determinant : " & dblDeterminant, vbOKOnly + vbInformation, Me.Caption
            
            blnNonSingular = mdlProcedures.IsNonSingular(dblSource, intMatrixMode)
            
            Me.txtNote.Text = ""
            
            If Not blnNonSingular Then
                Me.txtNote.Text = Me.txtNote.Text & "Matriks Singular"
            End If
            
            dblResult = mdlProcedures.SetInverse(dblSource, intMatrixMode)
            
            If dblDeterminant = 0# Then
                mdlProcedures.SetMatrixToFlexGrid Me.flxInverse, dblResult, intMatrixMode, True
            Else
                mdlProcedures.SetMatrixToFlexGrid Me.flxInverse, dblResult, intMatrixMode
            End If
            
            blnSquare = mdlProcedures.IsSquare(dblSource, intMatrixMode)
            
            If Not blnSquare Then
                If Not Trim(Me.txtNote.Text) = "" Then Me.txtNote.Text = Me.txtNote.Text & ", "
                
                Me.txtNote.Text = Me.txtNote.Text & "Bukan Matriks Square"
            End If
            
            blnSymmetric = mdlProcedures.IsSymmetric(dblSource, intMatrixMode)
            
            If Not blnSymmetric Then
                If Not Trim(Me.txtNote.Text) = "" Then Me.txtNote.Text = Me.txtNote.Text & ", "
                
                Me.txtNote.Text = Me.txtNote.Text & "Bukan Matriks Simetris"
            End If
            
            If blnNonSingular And blnSquare And blnSymmetric Then
                Me.txtNote.Text = "Semua Perhitungan Dapat Dijalankan"
            End If
        ElseIf objToolbarConst = PivotTl Then
            If mdlProcedures.IsUpperTringular(dblSource, intMatrixMode) Then Exit Sub

            mdlPivot.CalculatePivot dblSource, dblResult, intMatrixMode
            
            mdlProcedures.SetMatrixToFlexGrid Me.flxMatrix, dblResult, intMatrixMode
            
            blnNonSingular = False
            blnSquare = False
            blnSymmetric = False
        ElseIf objToolbarConst = EigenTl Then
            If Not blnNonSingular Then Exit Sub
            If Not blnSquare Then Exit Sub
            If Not blnSymmetric Then Exit Sub
            
            frmEigen.Show vbModal
        ElseIf objToolbarConst = LUTl Then
            If Not blnNonSingular Then Exit Sub
            If Not blnSquare Then Exit Sub
            
            frmLU.Show vbModal
        ElseIf objToolbarConst = SVDTl Then
            If Not blnNonSingular Then Exit Sub
            If Not blnSquare Then Exit Sub
            If Not blnSymmetric Then Exit Sub
            
            frmSVD.Show vbModal
        End If
    End If
End Sub

Private Sub ArrangeGrid(Optional ByVal intGrid As Integer = 0)
    Me.flxMatrix.Rows = intGrid
    Me.flxMatrix.Cols = intGrid
    
    Me.flxInverse.Rows = intGrid
    Me.flxInverse.Cols = intGrid
    
    Me.flxMatrix.RowHeightMin = Me.txtMatrix.Height
    Me.flxInverse.RowHeightMin = Me.txtMatrix.Height
    
    Dim lngCounter As Long
    
    For lngCounter = 0 To Me.flxMatrix.Cols - 1
        Me.flxMatrix.ColWidth(lngCounter) = 1300
        Me.flxInverse.ColWidth(lngCounter) = 1300
    Next lngCounter
End Sub

Private Sub FillTextInput()
    If Trim(Me.txtMatrix.Text) = "" Then Exit Sub
    
    If Not IsNumeric(Me.txtMatrix.Text) Then Me.txtMatrix.Text = "0"
    
    With Me.flxMatrix
        .TextMatrix(.Row, .Col) = Format(CCur(Me.txtMatrix.Text), "0.00")
    End With
End Sub

Private Sub ArrangeTextbox()
    With Me.flxMatrix
        Me.txtMatrix.Top = .Top + .CellTop
        Me.txtMatrix.Left = .Left + .CellLeft
        Me.txtMatrix.Width = .CellWidth
        Me.txtMatrix.Height = .CellHeight
        
        If Trim(.TextMatrix(.Row, .Col)) = "" Then
            Me.txtMatrix.Text = ""
        Else
            Me.txtMatrix.Text = .TextMatrix(.Row, .Col)
        End If
    End With
    
    Me.txtMatrix.Visible = True
    
    Me.txtMatrix.SetFocus
End Sub

Private Sub NewFile()
    If blnMatrix Then
        Dim intRow As Integer
        Dim intColumn As Integer
        
        For intRow = 0 To Me.flxMatrix.Rows - 1
            For intColumn = 0 To Me.flxMatrix.Cols - 1
                Me.flxMatrix.TextMatrix(intRow, intColumn) = ""
                Me.flxInverse.TextMatrix(intRow, intColumn) = ""
            Next intColumn
        Next intRow
        
        blnNonSingular = False
        blnSquare = False
        blnSymmetric = False
    End If
End Sub

Private Sub OpenFile()
    On Local Error GoTo ErrHandler
    
    With Me.cdlFile
        .InitDir = mdlGlobal.strPath
        .Filter = "Excel Files (*.xls)|*.xls"
        
        .ShowOpen
        
        If Not Trim(.FileName) = "" Then
            If Not mdlGlobal.fso.FileExists(.FileName) Then
                MsgBox "Data Tidak Ada", vbCritical, Me.Caption
            Else
                If Not blnMatrix Then
                    blnMatrix = True
                    
                    ShowMatrix
                End If
                
                Dim dblSource() As Double
                
                dblSource = mdlExcel.GetMatrix(.FileName)
                
                If mdlExcel.intMatrixChange = 2 Then
                    Me.cmbMatrix.ListIndex = 0
                Else
                    Me.cmbMatrix.ListIndex = mdlExcel.intMatrixChange - 3
                End If
                
                mdlProcedures.SetMatrixToFlexGrid Me.flxMatrix, dblSource, Me.cmbMatrix.ListIndex + 2
                
                mdlExcel.intMatrixChange = 0
                
                blnNonSingular = False
                blnSquare = False
                blnSymmetric = False
            End If
        End If
    End With
    
ErrHandler:
End Sub

Private Sub SaveFile()
    On Local Error GoTo ErrHandler
    
    If Not blnMatrix Then Exit Sub
    
    With Me.cdlFile
        .InitDir = mdlGlobal.strPath
        .Filter = "Excel Files (*.xls)|*.xls"
        
        .ShowSave
        
        If Trim(.FileName) = "" Then Exit Sub
        
        Dim dblSource() As Double
    
        dblSource = Me.SourceMatrix
        
        mdlExcel.SaveMatrix dblSource, .FileName, Me.cmbMatrix.ListIndex + 2
    End With
    
ErrHandler:
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get SourceMatrix() As Double()
    SourceMatrix = mdlProcedures.InputMatrix(Me.flxMatrix, Me.flxMatrix.Rows - 1, Me.flxMatrix.Cols - 1)
End Property

Public Property Get InverseMatrix() As Double()
    InverseMatrix = mdlProcedures.InputMatrix(Me.flxInverse, Me.flxInverse.Rows - 1, Me.flxInverse.Cols - 1)
End Property

Public Property Get MatrixMode() As Integer
    MatrixMode = Me.cmbMatrix.ListIndex + 2
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
End Property

Public Property Let MatrixMode(ByVal intMatrixMode As Integer)
    Me.cmbMatrix.ListIndex = intMatrixMode - 2
    
    blnMatrix = True
    
    ShowMatrix True
    
    blnNonSingular = False
    blnSquare = False
    blnSymmetric = False
End Property
