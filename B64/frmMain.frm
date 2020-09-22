VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base64 Codec"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   27
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdViewDecoded 
      Caption         =   "View Decoded"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   26
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton cmdViewEncoded 
      Caption         =   "View Encoded"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1880
      Left            =   4440
      TabIndex        =   12
      Top             =   60
      Width           =   2655
      Begin VB.Label lblSave 
         Alignment       =   2  'Center
         Caption         =   "n/a"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblDecode 
         Alignment       =   2  'Center
         Caption         =   "n/a"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblEncode 
         Alignment       =   2  'Center
         Caption         =   "n/a"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblLoad 
         Alignment       =   2  'Center
         Caption         =   "n/a"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "MB / Sec"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "MB / Sec"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "MB / Sec"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "MB / Sec"
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Save :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Decode :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Encode :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Load :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decode"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Save To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   120
      TabIndex        =   4
      Top             =   760
      Width           =   4215
      Begin VB.TextBox txtSave 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   2520
      End
      Begin VB.CommandButton cmdSaveBrowse 
         Caption         =   "..."
         Height          =   315
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Filename :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog dlgLoadSave 
      Left            =   4680
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Load From"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4215
      Begin VB.CommandButton cmdLoadBrowse 
         Caption         =   "..."
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtLoad 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   2520
      End
      Begin VB.Label Label1 
         Caption         =   "Filename :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()

frmAbout.Show vbModal

End Sub

Private Sub cmdDecode_Click()

cmdDecode.Caption = "Wait"
LockAllButton

StartTime = GetTickCount
clsBase64.Decode Buffer2, Buffer1
StopTime = GetTickCount

Elapsed = (StopTime - StartTime) / 1024
lblDecode.Caption = Str(Round((UBound(Buffer1) + 1) / 1048576 / Elapsed, 3))

UnlockAllButton
cmdDecode.Caption = "Decode"

End Sub

Private Sub cmdEncode_Click()

cmdEncode.Caption = "Wait"
LockAllButton

StartTime = GetTickCount
clsBase64.Encode Buffer1, Buffer2
StopTime = GetTickCount

Elapsed = (StopTime - StartTime) / 1024
lblEncode.Caption = Str(Round((UBound(Buffer1) + 1) / 1048576 / Elapsed, 3))

UnlockAllButton
cmdEncode.Caption = "Encode"

End Sub

Private Sub cmdLoad_Click()

cmdLoad.Caption = "Wait"
LockAllButton

StartTime = GetTickCount
clsBase64.Load txtLoad.Text, Buffer1
StopTime = GetTickCount

Elapsed = (StopTime - StartTime) / 1024
lblLoad.Caption = Str(Round((UBound(Buffer1) + 1) / 1048576 / Elapsed, 3))

UnlockAllButton
cmdLoad.Caption = "Load"

End Sub

Private Sub cmdSave_Click()

cmdSave.Caption = "Wait"
LockAllButton

StartTime = GetTickCount
clsBase64.Save Buffer1, txtSave.Text
StopTime = GetTickCount

Elapsed = (StopTime - StartTime) / 1024
lblSave.Caption = Str(Round((UBound(Buffer1) + 1) / 1048576 / Elapsed, 3))

UnlockAllButton
cmdSave.Caption = "Save"

End Sub

Private Sub cmdLoadBrowse_Click()

dlgLoadSave.ShowOpen
txtLoad.Text = dlgLoadSave.FileName

End Sub

Private Sub cmdSaveBrowse_Click()

dlgLoadSave.ShowSave
txtSave.Text = dlgLoadSave.FileName

End Sub

Private Sub cmdViewDecoded_Click()

cmdViewDecoded.Caption = "Wait"
LockAllButton

clsBase64.AryToStr Buffer1, tmpContent
frmDecoded.txtDecoded.Text = tmpContent
frmDecoded.Show

UnlockAllButton
cmdViewDecoded.Caption = "View Decoded"

End Sub

Private Sub cmdViewEncoded_Click()

cmdViewEncoded.Caption = "Wait"
LockAllButton

clsBase64.AryToStr Buffer2, tmpContent
frmEncoded.txtEncoded.Text = tmpContent
frmEncoded.Show

UnlockAllButton
cmdViewEncoded.Caption = "View Encoded"

End Sub

Private Sub LockAllButton()

cmdLoad.Enabled = False
cmdSave.Enabled = False
cmdEncode.Enabled = False
cmdDecode.Enabled = False
cmdViewEncoded.Enabled = False
cmdViewDecoded.Enabled = False
End Sub

Private Sub UnlockAllButton()

cmdLoad.Enabled = True
cmdSave.Enabled = True
cmdEncode.Enabled = True
cmdDecode.Enabled = True
cmdViewEncoded.Enabled = True
cmdViewDecoded.Enabled = True

End Sub

