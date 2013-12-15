VERSION 5.00
Begin VB.Form frmPop 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2160
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMove 
      Interval        =   5
      Left            =   360
      Top             =   1440
   End
   Begin VB.PictureBox picClick 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   360
      ScaleHeight     =   960
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   360
      Width           =   1815
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "has just signed in"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   480
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'intMove = 1 Then form is moving up
'intMove = 2 Then form is movin down
Private intMove As Integer



Private Sub Form_Click()

    picClick_Click

End Sub

Private Sub Form_Load()

    Dim intLen As Integer
    
    Me.Top = Screen.Height + 5
    Me.Left = Screen.Width - (Me.Width + 90)
    intLen = Len(lblName.Caption) / 2

    On Error Resume Next
    lblName.Move picClick.Width / intLen, picClick.Height / 4
    lblInfo.Move picClick.Width / 6, picClick.Height / 2
    intMove = 1

End Sub

Private Sub lblInfo_Click()

    picClick_Click

End Sub

Private Sub lblName_Click()

    picClick_Click

End Sub

Private Sub picClick_Click()

    strSendMessagePopup = Me.Tag
    frmOnline.XMSNC1.MSNSendMessage

End Sub

Private Sub tmrMove_Timer()

    tmrMove.Interval = 5

    If intMove = 1 Then
    
        Me.Top = (Me.Top - 30)
    
        If (Me.Top <= Screen.Height - (Me.Height - 10)) Then
      
            tmrMove.Interval = 6000
            intMove = 2
      
        End If
  
    ElseIf intMove = 2 Then
      
        Me.Top = (Me.Top + 30)
    
        If Me.Top >= (Screen.Height + 10) Then
      
            tmrMove.Enabled = False
            Unload Me
    
        End If
    
    End If

End Sub
