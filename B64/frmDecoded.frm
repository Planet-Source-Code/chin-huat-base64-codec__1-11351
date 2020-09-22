VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDecoded 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Decoded"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtDecoded 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6800
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDecoded.frx":0000
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
Attribute VB_Name = "frmDecoded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

With txtDecoded
    .Width = Me.Width - 120 * 3
    .Height = Me.Height - 120 * 5
End With

End Sub
