VERSION 5.00
Begin VB.Form frmPasswordXP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   Icon            =   "frmPasswordXP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPasswordXP.frx":030A
   ScaleHeight     =   4080
   ScaleWidth      =   6270
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
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   3375
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
      Left            =   1320
      TabIndex        =   1
      Top             =   2570
      Width           =   3615
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
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "frmPasswordXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
