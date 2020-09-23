VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Picture In System Properties"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   Icon            =   "SystemChanger.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Model 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   17
      Text            =   "Model Title = Windows is Funny"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Windows 98/95/3.1"
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   1920
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Window XP/2000/NT"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   2715
      TabIndex        =   7
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   735
      Left            =   4200
      Picture         =   "SystemChanger.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load"
      Height          =   735
      Left            =   3360
      Picture         =   "SystemChanger.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   855
      Left            =   4320
      Picture         =   "SystemChanger.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   855
      Left            =   2280
      Picture         =   "SystemChanger.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply all settings."
      Height          =   855
      Left            =   120
      Picture         =   "SystemChanger.frx":2FF2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Image Files|*.jpg;*.bmp;*.gif;*.ico;*.cur"
   End
   Begin VB.Image Image1 
      Height          =   1890
      Left            =   3000
      Picture         =   "SystemChanger.frx":3CBC
      Top             =   3480
      Width           =   2865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   2280
      Width           =   45
   End
   Begin VB.Label Label1 
      Caption         =   "Max 200 x 180"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y
Dim test2 As Boolean
Public test As Boolean
Private Sub Command1_Click()

Dim FileName, FilePic As String
'check it
If Option1.Value = False And Option2.Value = False Then Call ErrorOS

If test = False Then
GoTo 2:
ElseIf test2 = True Then
Call Error
End If
'Yes we made it once again
2
ret = MsgBox("This will save all, are you sure you want to save", vbYesNo)
If ret = vbYes Then
On Error Resume Next
'Kill ("C:\WINDOWS\system32\oeminfo.ini")
'Kill ("C:\WINDOWS\system32\OEMlogo.bmp")
MsgBox ("Save was a sucess press Windows Key+Pause Break to see changes"), vbInformation
If Option1.Value = True And Option2.Value = False Then
'if you are a XP/2000/NT user then the file will be saved to a diffeant location
FileName = "C:\WINDOWS\system32\oeminfo.ini"
FilePic = "C:\WINDOWS\system32\OEMlogo.bmp"
'This is the .ini file
Open FileName For Output As #1
Print #1, "[General]"
Print #1, "Manufacturer=Home Built PC"
Print #1, "Model=" & Model
Print #1, ""
Print #1, "[Support Information]"
Print #1, "Line1=" & Text1.Text
Print #1, "Line2=" & Text2.Text
Print #1, "Line3=" & Text3.Text
Print #1, "Line4=" & Text4.Text
Print #1, "Line5=" & Text5.Text
Print #1, "Line6=" & Text6.Text
Print #1, "Line7=" & Text5.Text
Close #1
SavePicture Picture1.picture, FilePic

'
'
'
ElseIf Option2.Value = True And Option1.Value = False Then
'if you are a 98/95/3.1 user then the file will be saved to a diffeant location
FileName = "C:\Windows\system\oeminfo.ini"
FilePic = "C:\Windows\system\OEMlogo.bmp"
'This is the .ini file
Open FileName For Output As #1
Print #1, "[General]"
Print #1, "Manufacturer=Home Built PC"
Print #1, "Model=" & Model
Print #1, ""
Print #1, "[Support Information]"
Print #1, "Line1=" & Text1.Text
Print #1, "Line2=" & Text2.Text
Print #1, "Line3=" & Text3.Text
Print #1, "Line4=" & Text4.Text
Print #1, "Line5=" & Text5.Text
Print #1, "Line6=" & Text6.Text
Print #1, "Line7=" & Text5.Text
Close #1
'End
ElseIf ret = vbNo Then
End If
End If
End Sub
Sub Error()
MsgBox ("It appears that your picture is to big" & vbNewLine & "Please select another one"), vbExclamation
End Sub
Sub ErrorOS()
MsgBox ("Please select an OS then click apply."), vbInformation
End Sub
Private Sub Command2_Click()
Unload Form1
End
End Sub

Private Sub Command3_Click()
MsgBox ("Yar System Properties Hack Version 1.0" & vbNewLine & "                    Joshua Nixon"), vbInformation
End Sub

Private Sub Command4_Click()

x = 0
y = 0
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
Picture1.picture = LoadPicture(CommonDialog1.FileName)
'create an autofit
'Check for errors
 Picture1.AutoSize = True
    x = ScaleX(Picture1.Width, vbTwips, vbPixels)
    y = ScaleY(Picture1.Height, vbTwips, vbPixels)
    sizeof = x & "," & y
    Label2.Caption = "Dimensions: " & sizeof
 
 If y > 180 Then
    MsgBox ("The height of the picture is to large," & vbNewLine & "select another the less than or equal to 180"), vbExclamation
 GoTo 7:
 ElseIf x > 200 Then
    MsgBox ("The width of the picture is to large," & vbNewLine & "select another the less than or equal to 180"), vbExclamation
GoTo 7:
ElseIf x > 200 And y < 200 Then
    MsgBox ("The width and height of the picture is to large," & vbNewLine & "select another the less than or equal to 180"), vbExclamation
GoTo 7:
Text = True
End If
' Yes it passed so lets go on
If x > 0 And x <= 200 And y > 0 And y <= 180 Then
    Autofit Picture1
 test = False
 End If
'error check
7:
Picture1.Height = 2655
Picture1.Width = 2775
Picture1.AutoRedraw = False
Autofit Picture1
test = True
End Sub

Function Autofit(picture As PictureBox)
On Error Resume Next
picture.PaintPicture picture, 0, 0, picture.ScaleWidth, picture.ScaleHeight
End Function
Private Sub Command5_Click()
Picture1.picture = LoadPicture()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

