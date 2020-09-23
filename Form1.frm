VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddIn 
   Caption         =   "S.AE API Viewer 6.0"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkFormatApi 
      Caption         =   "Insert with formatted line breaks"
      Height          =   200
      Left            =   7440
      TabIndex        =   21
      Top             =   360
      Value           =   1  'Checked
      Width           =   2595
   End
   Begin VB.CheckBox chkComments 
      Caption         =   "Insert API comment separators"
      Height          =   200
      Left            =   7440
      TabIndex        =   20
      Top             =   0
      Width           =   2595
   End
   Begin VB.ListBox Declares 
      Height          =   840
      ItemData        =   "Form1.frx":0442
      Left            =   9000
      List            =   "Form1.frx":0444
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox Types 
      Height          =   840
      ItemData        =   "Form1.frx":0446
      Left            =   9480
      List            =   "Form1.frx":0448
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox Constants 
      Height          =   840
      ItemData        =   "Form1.frx":044A
      Left            =   9960
      List            =   "Form1.frx":044C
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog ad 
      Left            =   6480
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox ado 
      Height          =   480
      Left            =   7800
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   16
      Top             =   3720
      Width           =   1200
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "I&nsert"
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Declare Scope"
      Height          =   975
      Left            =   5760
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
      Begin VB.OptionButton optPrivate 
         Caption         =   "Private"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optPublic 
         Caption         =   "Public"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtAPICode 
      Height          =   1455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   4560
      Width           =   5775
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "Form1.frx":044E
      Left            =   0
      List            =   "Form1.frx":0450
      TabIndex        =   5
      Top             =   1560
      Width           =   5655
   End
   Begin VB.TextBox txtSearchList 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   6855
   End
   Begin VB.ComboBox cboAPITypes 
      Height          =   315
      ItemData        =   "Form1.frx":0452
      Left            =   0
      List            =   "Form1.frx":0454
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Items:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4290
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Items:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Type The First Few Letters Of The Word You are looking For:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "API Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Op 
         Caption         =   "Open File"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit
Private decs(5000) As String '  TRUE DECLARES
Private typs(5000) As String '  TRUE TYPES
Private conts(5000) As String
Private intype As Boolean
Private tsplit() As String
Private Declare Function SendMessage Lib _
      "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
      ByVal wMsg As Long, _
      ByVal wParam As Long, _
      lParam As Any) As Long

Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2
Dim CurrentFile As String
Dim i As Long
Dim tmp As String
Private Sub cboAPITypes_Click()
     List1.Clear

     If cboAPITypes.text = "Declares" Then

          For i = 0 To Declares.ListCount
               If Declares.List(i) = "" Then GoTo 1
               List1.AddItem Declares.List(i)
1
          Next i

     End If

     If cboAPITypes.text = "Constants" Then

          For i = 0 To Constants.ListCount
               If Constants.List(i) = "" Then GoTo 2
               List1.AddItem Constants.List(i)
2
          Next i

     End If
     If cboAPITypes.text = "Types" Then

          For i = 0 To Types.ListCount
               If Types.List(i) = "" Then GoTo 3
               List1.AddItem Types.List(i)
3
          Next i
     End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.RemoveItem List1.text
End Sub
Private Sub Command3_Click()
txtAPICode.text = ""
End Sub

Private Sub Command5_Click()
 Clipboard.Clear
     Clipboard.SetText txtAPICode.text
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Op_Click()
On Error Resume Next
ad.Filter = "API Text File|*.Txt"
ad.ShowOpen
   cboAPITypes.text = "Declares"
     cboAPITypes.AddItem "Declares"
     cboAPITypes.AddItem "Types"
     cboAPITypes.AddItem "Constants"
     
     Dim tmp As String
CurrentFile = ad.FileName
If Dir(CurrentFile) > "" Then
     Open CurrentFile For Input As #1
     Do Until EOF(1)
          Line Input #1, tmp
          check tmp
     Loop

     Close #1
     End If
End Sub
Public Sub ctype(text As String)

     If LCase(text) = "end type" Then intype = False: text = " " & text
     typs(Types.ListCount) = vbCrLf & typs(Types.ListCount) & text & vbCrLf

End Sub
Public Sub check(text As String)
     If intype = True Then ctype (text): Exit Sub
     tsplit() = Split(text, " ")
     tmp = LCase(Left(text, 4))
     If tmp = "decl" Then
          Declares.AddItem tsplit(2)
          decs(Declares.ListCount) = text
     End If
     If tmp = "type" Then
          intype = True
          Types.AddItem Right(text, Len(text) - 5)
          typs(Types.ListCount) = text
          typs(Types.ListCount) = typs(Types.ListCount) & vbCrLf
     End If
     If tmp = "cons" Then
          Constants.AddItem Right(text, Len(text) - 6)
     End If
End Sub
Public Sub fdec(fake As String)
On Error Resume Next
Dim ApiStuff As String
Dim OldLength As Long
Dim ChrCount As Long
Dim ScopeStyle As String
Dim FinalScope As String
Dim sListText As String

sListText = List1.text
sListText = Trim(sListText)
OldLength = Len(sListText)
ChrCount = 89 - OldLength - 15

For i = 1 To Declares.ListCount
     tsplit() = Split(decs(i), " ")

     If LCase(tsplit(2)) = LCase(fake) Then

          If optPublic.Value = True Then ScopeStyle = "Public "
          If optPrivate.Value = True Then ScopeStyle = "Private "

               If chkFormatApi.Value = vbChecked Then
                    ApiStuff = Replace(decs(i), "Lib", "Lib _ " & vbCrLf & "     ")
                    ApiStuff = Replace(ApiStuff, ",", ", _ " & vbCrLf & "     ")
                    FinalScope = ScopeStyle & ApiStuff & vbCrLf
               Else
                    FinalScope = ScopeStyle & decs(i) & vbCrLf
               End If

          If chkComments = vbChecked Then
               FinalScope = "'  API'S FOR " & UCase(sListText) & ": " & String(ChrCount, "=") & _
               " STARTS HERE" & vbCrLf & vbCrLf _
               & FinalScope & vbCrLf _
               & "'  API'S FOR " & UCase(sListText) & ": " & String(ChrCount, "=") & _
               " ENDS HERE" & vbCrLf
          End If

     txtAPICode.text = txtAPICode.text & vbCrLf & FinalScope
     End If
Next i
End Sub
Private Sub cboAPITypes_LostFocus()
     cboAPITypes.SelLength = 0
End Sub
Public Sub ftyp(fake As String)
     For i = 1 To Types.ListCount
          tsplit() = Split(typs(i), " ")
          If LCase(tsplit(1)) = LCase(fake) & vbCrLf Then
               If optPublic.Value = True Then txtAPICode.text = txtAPICode.text & "Public "
               If optPrivate.Value = True Then txtAPICode.text = txtAPICode.text & "Private "
               Do Until LCase(Left(typs(i), 1)) = "t"
                    typs(i) = Right(typs(i), Len(typs(i)) - 1)
               Loop
               txtAPICode.text = txtAPICode.text & typs(i) & vbCrLf
          End If
     Next i
End Sub
Private Sub List1_DblClick()
txtAPICode.text = ""
     If cboAPITypes.text = "Declares" Then
          fdec (List1.List(List1.ListIndex))
     End If

     If cboAPITypes.text = "Constants" Then
          If optPublic.Value = True Then txtAPICode.text = txtAPICode.text & "Public Const " & List1.List(List1.ListIndex) & vbCrLf
          If optPrivate.Value = True Then txtAPICode.text = txtAPICode.text & "Private Const " & List1.List(List1.ListIndex) & vbCrLf
          txtAPICode.text = txtAPICode.text '& vbCrLf
     End If

     If cboAPITypes.text = "Types" Then
          ftyp (List1.List(List1.ListIndex))
     End If
End Sub
Private Sub txtSearchList_Change()
List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, 0&, ByVal (txtSearchList.text))
End Sub
