VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form PriceCompareFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price Compare"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser priceBrowser 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      ExtentX         =   11245
      ExtentY         =   16536
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "PriceCompareFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub GetPriceComparisons(product As String)
    
End Sub
