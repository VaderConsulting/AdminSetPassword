VERSION 5.00
Begin VB.Form frmPasswordXPClassic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "frmPasswordXPClassic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPasswordXPClassic.frx":030A
   ScaleHeight     =   3255
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblDomain 
      BackStyle       =   0  'Transparent
      Caption         =   "Domain"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2595
      TabIndex        =   2
      Top             =   2000
      Width           =   2415
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2595
      TabIndex        =   1
      Top             =   1600
      Width           =   2655
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2595
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
End
Attribute VB_Name = "frmPasswordXPClassic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
