VERSION 5.00
Begin VB.Form frmPasswordEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Password"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmPasswordEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label lblRetry 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Image imgFeedback 
      Height          =   480
      Left            =   240
      Picture         =   "frmPasswordEntry.frx":030A
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblPasswordHint 
      Alignment       =   2  'Center
      Caption         =   "This application is password protected.  Enter the correct password to continue."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmPasswordEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCanUnload As Boolean
Dim intRetryCount As Integer

Private Sub cmdOK_Click()
    If frmPasswordEntry.txtPassword.Text = strApplicationPasswordCleartext Then
        bCanUnload = True
        Unload Me
    Else
        intRetryCount = intRetryCount + 1
        lblRetry = "Incorrect password.  " & 3 - intRetryCount & " chances remaining."
        frmPasswordEntry.imgFeedback.Visible = True
        
        frmPasswordEntry.txtPassword.Text = ""
        frmPasswordEntry.txtPassword.SetFocus
        
        If intRetryCount = 3 Then
            bCanUnload = True
            End
        End If
    End If
End Sub

Private Sub Form_Load()
    bCanUnload = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not bCanUnload Then Cancel = 1
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Capture carriage return
    If KeyCode = 13 Then cmdOK_Click
End Sub
