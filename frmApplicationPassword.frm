VERSION 5.00
Begin VB.Form frmApplicationPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Application Password"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "frmApplicationPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblSetPasswordHint 
      Caption         =   "Enter the password that is used to control access to this application."
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
      Width           =   4335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000015&
      X1              =   120
      X2              =   4440
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   4440
      Y1              =   1095
      Y2              =   1095
   End
End
Attribute VB_Name = "frmApplicationPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    ' Before saving the password, we have to encrypt it.
    strApplicationPasswordCleartext = Me.txtPassword.Text
    strApplicationPasswordEncrypted = Encrypt(strApplicationPasswordCleartext)
    
    SaveIniValue "Setup", "Password", strApplicationPasswordEncrypted, strAppPath & "setup.ini"
    Unload Me
End Sub
