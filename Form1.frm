VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INI Database"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "New"
      Height          =   345
      Left            =   1770
      TabIndex        =   6
      Top             =   3465
      Width           =   795
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2895
      Left            =   2655
      TabIndex        =   5
      Top             =   465
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   5106
      _Version        =   393217
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   4965
      TabIndex        =   4
      Top             =   45
      Width           =   1950
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   60
      Width           =   2235
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deleite"
      Height          =   300
      Left            =   75
      TabIndex        =   2
      Top             =   3465
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   300
      Left            =   4485
      TabIndex        =   1
      Top             =   3435
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   2550
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command3_Click()
Dim x As Integer
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then List1.Selected(x) = False
Next x
Text1.Text = ""
Text2.Text = ""
RichTextBox1.Text = ""
Command1.Enabled = True
Command1.Caption = "Add"
End Sub
Private Sub Form_Load()
Command1.Enabled = False
PopulateEntrys
End Sub
Private Sub Command1_Click()
Dim x As Integer
If Command1.Caption = "Add" Then
For x = 0 To List1.ListCount - 1
If Text1.Text = List1.List(x) Then Exit Sub
Next x
SetInitEntry Text1.Text, "Name", Text2.Text, App.Path & "\MyINI.ini"
SetInitEntry Text1.Text, "Note", Replace(RichTextBox1.Text, vbCrLf, "***RETURN***"), App.Path & "\MyINI.ini"
PopulateEntrys
Command1.Caption = "Update"
ElseIf Command1.Caption = "Update" Then
For x = 0 To List1.ListCount - 1
If Text1.Text = List1 Then
SetInitEntry Text1.Text, "Name", Text2.Text, App.Path & "\MyINI.ini"
SetInitEntry Text1.Text, "Note", Replace(RichTextBox1.Text, vbCrLf, "***RETURN***"), App.Path & "\MyINI.ini"
Exit Sub
Else
MsgBox "Description has changed ! You have to press New to add a new entree !"
Text1.Text = List1
Exit Sub
End If
Next x
End If
End Sub
Private Sub Command2_Click()
Dim x As Integer
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
SetInitEntry List1.List(x), , , App.Path & "\MyINI.ini"
Text1.Text = ""
Text2.Text = ""
RichTextBox1.Text = ""
End If
Next x
PopulateEntrys
End Sub
Private Sub PopulateEntrys()
Dim sParts() As String
Dim i As Integer
List1.Clear
sParts() = Split(GetInitEntry(vbNullString, "", "", App.Path & "\MyINI.ini"), Chr(0))
For i = 0 To UBound(sParts) - 1
List1.AddItem sParts(i)
Next i
End Sub
Private Sub List1_Click()
Dim x As Integer
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
Text2.Text = GetInitEntry(List1.List(x), "Name", "", App.Path & "\MyINI.ini")
RichTextBox1.Text = Replace(GetInitEntry(List1.List(x), "Note", "", App.Path & "\MyINI.ini"), "***RETURN***", vbCrLf)
Text1.Text = List1.List(x)
Exit For
End If
Next x
Command1.Enabled = True
Command1.Caption = "Update"
End Sub
