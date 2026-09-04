VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Set Password"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   3720
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3720
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUsernames 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Line lneBorderWhite 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   7200
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Line lneBorderGrey 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   7200
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblUsername 
      Alignment       =   2  'Center
      Caption         =   "Usernames"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileLoadUserlist 
         Caption         =   "Load Userlist"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPasswordS 
      Caption         =   "Passwords"
      Begin VB.Menu mnuPasswordsSetApplication 
         Caption         =   "Set Application"
      End
      Begin VB.Menu mnuPasswordsSetStandard 
         Caption         =   "Set Standard"
      End
      Begin VB.Menu mnuPasswordsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPasswordsRandomise 
         Caption         =   "Randomise"
      End
      Begin VB.Menu mnuPasswordsStandardise 
         Caption         =   "Standardise"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit

Private Sub Form_Load()
    Dim bRequirePassword As Boolean
    
    If Right(App.Path, 1) = "\" Then
        strAppPath = App.Path
    Else
        strAppPath = App.Path & "\"
    End If
    
    GetConfig
    
    strDefaultPasswordClearText = Decrypt(strDefaultPasswordEncrypted)
    Debug.Print "Standard Password: " & strDefaultPasswordClearText
    strApplicationPasswordCleartext = Decrypt(strApplicationPasswordEncrypted)
    Debug.Print "Application Password: " & strApplicationPasswordCleartext
    
    Me.Show
    
    If Len(strApplicationPasswordCleartext) > 0 Then
        frmPasswordEntry.Show vbModal
    End If
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub
    
Private Sub mnuFileLoadUserlist_Click()
    Dim strIn As String
    Dim strUserInfo() As String
    Dim i As Integer
    
    If Dir(strapplist & "userlist.txt") <> "" Then
        Open strAppPath & "userlist.txt" For Input As #1
            Do Until EOF(1)
                i = i + 1
                Line Input #1, strIn
                strUserInfo() = Split(strIn, ",", -1, vbTextCompare)
                ReDim Preserve oUsers(i)
                oUsers(i).Domain = strUserInfo(0) ' NOTE:  Even though OPTION BASE 1 is used, split always returns a 0 based array.
                oUsers(i).Username = strUserInfo(1)
                If UBound(strUserInfo()) = 2 Then
                    oUsers(i).EncryptedPassword = strUserInfo(2)
                End If
            Loop
        Close 1
    Else
        MsgBox strAppPath & "userlist.txt not found.", vbExclamation + vbOKOnly, "Message"
    End If
End Sub

Private Sub mnuPasswordsRandomise_Click()
    OS = W2KWS
    
    Select Case OS
        Case OSType.WNT4Server
        Case OSType.WNT4WS
        Case OSType.WNT4TS
        Case OSType.W2KServer
        Case OSType.W2KWS
            frmPassword2000WS.lblUsername = "Fred Bloggs"
            frmPassword2000WS.lblPassword = "12345678"
            frmPassword2000WS.lblDomain = "Here"
            frmPassword2000WS.Show
        Case OSType.W2KTS
        Case OSType.WXPClassic
        Case OSType.WXPDomain
        Case OSType.WXPWorkgroup
        Case OSType.W2003Server
    End Select
End Sub

Private Sub mnuPasswordsSetApplication_Click()
    frmApplicationPassword.Show vbModal
End Sub

Private Sub mnuPasswordsSetStandard_Click()
    frmSetPassword.Show vbModal
End Sub

Private Sub mnuPasswordsStandardise_Click()
    Dim i As Integer
    Dim s As String
    Dim u() As String
    Dim sUsername As String
    
    ' Get all usernames from textbox, and add to array.
    s = Replace(frmMain.txtUsernames, vbCrLf, "", 1, -1, vbTextCompare)
    s = Replace(s, vbCr, ",", 1, -1, vbTextCompare)
    s = Replace(s, vbLf, ",", 1, -1, vbTextCompare)
    s = Replace(s, vbTab, ",", 1, -1, vbTextCompare)
    s = Trim(s)
    s = Replace(s, ",,", ",")
    
    u() = Split(s, ",")
    ' Set Password for all users to standard.
    For i = 0 To UBound(u())
        sUsername = u(i)
        
    Next i
End Sub

