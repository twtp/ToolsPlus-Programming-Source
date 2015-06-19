VERSION 5.00
Begin VB.Form AddWebPathsRouterBitTableEx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Web Paths - Router Bit Table Customization"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   10020
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6090
      TabIndex        =   4
      Text            =   "RBX_ID"
      Top             =   120
      Width           =   705
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   150
      TabIndex        =   3
      Top             =   1350
      Width           =   1875
   End
   Begin VB.TextBox WebPathName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4875
   End
   Begin VB.TextBox WebPathID 
      Height          =   285
      Left            =   5280
      TabIndex        =   1
      Text            =   "WP_ID"
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Router Bit Table Customization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "AddWebPathsRouterBitTableEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

