VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl clock 
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1965
   ScaleHeight     =   1590
   ScaleWidth      =   1965
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   1590
      Left            =   -45
      TabIndex        =   0
      Top             =   0
      Width           =   1995
      ExtentX         =   3519
      ExtentY         =   2805
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
Attribute VB_Name = "clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=============================================================================================
Private Sub UserControl_Resize()
  
  Width = 2000
  Height = 2100
  WB1.Move -30, -30, Width + 350, Height + 60
  
End Sub
'=============================================================================================
Private Sub UserControl_Show()
 
   Call loadDLL_webpage(WB1, App.Path & "\makedll.dll", 101)
 
End Sub
'=============================================================================================
Sub loadDLL_webpage(your_webbrowser As WebBrowser, _
                     dll_path As String, _
                     resource_number As Long, _
                     Optional custom_resource_name As String = "CUSTOM")
   Dim surl$
  
   surl = _
          "res://" & _
          dll_path$ & _
          "/" & custom_resource_name & _
          "/" & resource_number
 
    your_webbrowser.Silent = True
    your_webbrowser.Navigate surl
  
End Sub
'=============================================================================================
