VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "QuickC"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Functions 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      ItemData        =   "Form1.frx":259E8
      Left            =   3000
      List            =   "Form1.frx":25A07
      TabIndex        =   13
      Text            =   "Fuctions"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox Keywords 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      ItemData        =   "Form1.frx":25A6B
      Left            =   1560
      List            =   "Form1.frx":25AD2
      TabIndex        =   12
      Text            =   "Keywords"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox Headers 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "Form1.frx":25BE4
      Left            =   120
      List            =   "Form1.frx":25BF7
      TabIndex        =   11
      Text            =   "Headers"
      Top             =   960
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5520
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   700
      Left            =   5160
      Top             =   1440
   End
   Begin RichTextLib.RichTextBox Text 
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9340
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":25C20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label ENTBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "{...}"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4440
      TabIndex        =   14
      Top             =   960
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   255
      Left            =   5760
      Shape           =   5  'Rounded Square
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label OpenBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label SaveBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Label ClearBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label MinBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label CloseBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5760
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Label DateAndTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Date and Time"
      Top             =   6840
      Width           =   4515
   End
   Begin VB.Label CharInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Character count"
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QuickC"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      Height          =   7215
      Left            =   0
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Mover 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C0C0&
      X1              =   0
      X2              =   6120
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CHX, CHY As Integer



Public Function SaveFile()

    With CommonDialog1
            .DialogTitle = "Save - QuickNote"
            .CancelError = False
            .Filter = "C Source Code (*.c)|*.c|All Files (*.*)|*.*"
           .ShowSave
           
            If Len(.FileName) = 0 Then
        Return
            End If
                sfile = .FileName
            End With
        
     extension = LCase(Right(sfile, 3))
      
    
   
       Open sfile For Output As #1
        Print #1, Text.Text
        Close #1
            
       
End Function

Public Function OpenFile()

       With CommonDialog1
        .DialogTitle = "Open - QuickNote"
        .CancelError = False
                .Filter = "Text Files (*.txt)|*.txt|Rich Text Format(*.rtf)|*.rtf|All Files(*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
           Return
        End If
        sfile = .FileName
    End With
 Text.LoadFile sfile
   
       
End Function


Private Sub ENTBtn_Click()
Text.SelText = "{" & Chr$(13) + Chr$(10) & Chr(9) & Chr$(13) + Chr$(10) & "}" & Chr$(13) + Chr$(10)
End Sub

Private Sub ENTBtn_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ENTBtn.BackColor = RGB(0, 205, 0)
End Sub

Private Sub ENTBtn_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
ENTBtn.BackColor = &HFF00&
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 19 Then SaveFile
If KeyAscii = 15 Then OpenFile
End Sub

Private Sub Form_Load()
Me.Hide
Splash.Show
Me.Label2.ForeColor = RGB(0, 0, 200)

End Sub

Private Sub Fuctions_Click()
Text.Text = Text.Text & Functions.Text
End Sub

Private Sub Functions_Click()

Text.SelColor = &H8000&

Text.SelText = Functions.Text
Text.SelColor = 0
End Sub

Private Sub Headers_Click()

Text.SelColor = &HFF0000
Text.SelBold = True
Text.SelText = "#include <" & Headers.Text & ".h>" & Chr$(13) + Chr$(10)
Text.SelColor = &H0&
Text.SelBold = False
End Sub

Private Sub Keywords_Click()

Text.SelColor = &HC0&

Text.SelText = Keywords.Text
Text.SelColor = 0
End Sub

Private Sub Label1_Click()
About.Show
End Sub


Private Sub Mover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 On Error Resume Next
 If Button = 1 Then
        Me.Left = Me.Left + x
        Me.Top = Me.Top + y
    End If
End Sub

Private Sub CloseBTN_Click()
Closewindow.Show
End Sub

Private Sub CloseBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
CloseBTN.BorderStyle = 1
CloseBTN.BackColor = &H80&
CloseBTN.ForeColor = RGB(200, 250, 255)
End Sub

Private Sub CloseBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
CloseBTN.BorderStyle = 0
CloseBTN.BackColor = RGB(255, 20, 20)
CloseBTN.ForeColor = RGB(0, 20, 250)
End Sub

Private Sub MinBTN_Click()
Me.WindowState = 1
End Sub
Private Sub MinBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MinBTN.BackColor = RGB(50, 50, 255)
MinBTN.ForeColor = &HFFFF&
End Sub

Private Sub MinBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MinBTN.BackColor = &HFFFF00
MinBTN.ForeColor = &HC00000
End Sub
Private Sub ClearBTN_Click()
Text.Text = ""
End Sub

Private Sub ClearBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ClearBTN.BackColor = RGB(0, 205, 0)
End Sub

Private Sub ClearBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
ClearBTN.BackColor = &HFF00&
End Sub
Private Sub SaveBTN_Click()
On Error Resume Next
     SaveFile
End Sub

Private Sub SaveBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
SaveBTN.BackColor = RGB(0, 205, 0)
End Sub

Private Sub SaveBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
SaveBTN.BackColor = &HFF00&
End Sub
Private Sub OpenBTN_Click()
On Error Resume Next
    OpenFile
 End Sub
Private Sub OpenBTN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
OpenBTN.BackColor = RGB(0, 205, 0)
End Sub

Private Sub OpenBTN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
OpenBTN.BackColor = &HFF00&
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape1.BorderColor = RGB(255, 0, 0)
End Sub


Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape1.BorderColor = RGB(0, 25, 250)
End Sub


Private Sub Text_Change()
CharInfo.Caption = Len(Text.Text)
End Sub

Private Sub Text_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 19 Then SaveFile
If KeyAscii = 15 Then OpenFile
If KeyAscii = 9 Then Text.SelText = Chr(9)
CharInfo.Caption = Len(Text.Text)
End Sub


Private Sub Timer1_Timer()
DateAndTime.Caption = Date & "  " & Time
If Val(CharInfo.Caption) > 0 Then ClearBTN.Visible = True Else ClearBTN.Visible = False

End Sub

Private Sub Timer2_Timer()
If Label1.ForeColor = &H8000& Then
Label1.ForeColor = &H80&
Else
Label1.ForeColor = &H8000&
End If
End Sub
