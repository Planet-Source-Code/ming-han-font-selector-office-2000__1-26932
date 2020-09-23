VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Font 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox cmbfontname 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageCombo ic1 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   255
      BackColor       =   16777215
      ImageList       =   "iml"
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   4200
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "Font"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Click()
Public Event Change()
Dim Flags As String

Private Sub ic1_Click()
RaiseEvent Click
End Sub

Public Property Let Enabled(ena As Boolean)
ic1.Enabled = ena
End Property
Public Property Get Enabled() As Boolean
Enabled = ic1.Enabled
End Property

Public Property Let Text(str As String)
Flags = "Let"
ic1.ComboItems(str).Selected = True
Flags = ""
End Property

Public Property Get Text() As String
Text = ic1.SelectedItem.Key
End Property

Sub MakeFonts()

With cmbfontname
      For I = 0 To Screen.FontCount - 1
         .AddItem Screen.Fonts(I)
      Next I
      ' Set ListIndex to 0.
      .ListIndex = 0
   End With
On Error Resume Next
ic1.ComboItems.Clear

With ic1
pb1.Move .Left, .Top, .Width, .Height
End With

pb1.Visible = True
pb1.Max = cmbfontname.ListCount
pb1.Value = 0

iml.ListImages.Clear

For I = 0 To cmbfontname.ListCount - 1

    With img
        .CurrentX = 0
        .CurrentY = 0
        .FontSize = 11
        .FontItalic = False
        .FontBold = False
        .FontName = cmbfontname.List(I)
        .Width = ic1.Width
        .Height = ic1.Height
        .Cls
        img.Print cmbfontname.List(I)
        
    End With
    
    iml.ListImages.Add , , img.Image
    
    Set ic1.ImageList = iml
    ic1.ComboItems.Add , cmbfontname.List(I), , I + 1
    pb1.Value = I + 1
    
Next

pb1.Visible = False

End Sub

Private Sub ic1_Dropdown()
If ic1.ComboItems.Count = 0 Then Me.MakeFonts
End Sub

Private Sub ic1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Me.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
UserControl.Height = ic1.Height
ic1.Width = UserControl.Width
ic1.Move 0, 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Enabled", Me.Enabled, True
End Sub
