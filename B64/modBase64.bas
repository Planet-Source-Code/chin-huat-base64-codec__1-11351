Attribute VB_Name = "modBase64"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Buffer1() As Byte
Public Buffer2() As Byte
Public Span() As String
Public Unspan() As Byte
Public StartTime As Long
Public StopTime As Long
Public Elapsed As Single
Public tmpContent As String

Public clsBase64 As clsB64

Public Sub Main()

Set clsBase64 = New clsB64
clsBase64.Init

frmMain.Show

End Sub
