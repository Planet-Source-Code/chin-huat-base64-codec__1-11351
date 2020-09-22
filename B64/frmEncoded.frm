VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEncoded 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Encoded"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdUnspan 
      Caption         =   "Unspan"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSpan 
      Caption         =   "Span"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   1455
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
      TabIndex        =   1
      Top             =   60
      Width           =   4575
      Begin VB.CommandButton cmdSaveBrowse 
         Caption         =   "..."
         Height          =   315
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtSave 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "Filename :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   2400
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtEncoded 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6800
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmEncoded.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmEncoded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()

cmdSave.Caption = "Wait"
LockAllButton

txtEncoded.SaveFile txtSave.Text

UnlockAllButton
cmdSave.Caption = "Save"

End Sub

Private Sub cmdSpan_Click()

Dim Index As Long
Dim strTemp As String

cmdSpan.Caption = "Wait"
LockAllButton

clsBase64.Span Buffer2, Span, 76, vbCrLf

txtEncoded.Visible = False
txtEncoded.Text = ""

For Index = LBound(Span) To UBound(Span)
    strTemp = strTemp + Span(Index)
Next

txtEncoded.Text = strTemp
txtEncoded.Visible = True

UnlockAllButton
cmdSpan.Caption = "Span"

End Sub

Private Sub cmdUnspan_Click()

Dim Index As Long
Dim strTemp As String

cmdUnspan.Caption = "Wait"
LockAllButton

txtEncoded.Text = ""
txtEncoded.Visible = False

clsBase64.Unspan Span, Unspan, 76, vbCrLf
clsBase64.AryToStr Unspan, strTemp

txtEncoded.Text = strTemp
txtEncoded.Visible = True

UnlockAllButton
cmdUnspan.Caption = "Unspan"

End Sub

Private Sub Form_Resize()

With txtEncoded
    .Width = Me.Width - 120 * 3
    .Height = Me.Height - 120 * 5 - 1160
End With

End Sub

Private Sub LockAllButton()

cmdSave.Enabled = False
cmdSpan.Enabled = False
cmdUnspan.Enabled = False

End Sub

Private Sub UnlockAllButton()

cmdSave.Enabled = True
cmdSpan.Enabled = True
cmdUnspan.Enabled = True

End Sub
