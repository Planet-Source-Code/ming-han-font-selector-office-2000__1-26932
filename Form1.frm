VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin RichTextLib.RichTextBox rtf1 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin Project1.Font fs1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
fs1.MakeFonts
End Sub

Private Sub fs1_Click()
On Error Resume Next
rtf1.SelFontName = fs1.Text
End Sub

Private Sub rtf1_SelChange()
If Not IsNull(rtf1.SelFontName) Then
   fs1.Text = rtf1.SelFontName
End If
End Sub
