VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "dyr_workshop"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   960
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Harlow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer, s As Integer
Dim max As Boolean

Private Sub Timer1_Timer()
    If max = False And X <= 255 Then
        X = X + 3
    ElseIf max = True And X > 0 Then
        X = X - 3
    End If
    
    If X = 255 Then
        max = True
    ElseIf X = 0 Then
        max = False
    End If
    Label1.ForeColor = RGB(0, X, 0)
End Sub

Private Sub Timer2_Timer()
    s = s + 1
    Form1.Caption = Mid("dyr_workshop", 1, s)
    If s = 13 Then
        Form1.Caption = ""
        s = 0
    End If
End Sub
