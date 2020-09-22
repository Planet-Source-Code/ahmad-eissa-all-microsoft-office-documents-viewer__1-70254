VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Microsoft Office All Documents Viewer"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5115
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   9735
      ExtentX         =   17171
      ExtentY         =   9022
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
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8580
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oDocument As Object

Private Sub Command1_Click()
   Dim sFileName As String
   

   With CommonDialog1
      .FileName = ""
      .ShowOpen
      sFileName = .FileName
   End With
   

   If Len(sFileName) Then
      Set oDocument = Nothing
      WebBrowser1.Navigate sFileName
   End If
End Sub

Private Sub Command2_Click()
    WebBrowser1.Navigate "about:blank"
End Sub

Private Sub Form_Load()
   Command1.Caption = "Browse"
   With CommonDialog1
      .Filter = "Office Documents " & _
      "Office Documents (*.doc, *.xls, *.ppt)|*.doc;*.xls;*.ppt|Office 2007 Documents (*.docx, *.xlsx, *.pptx)|*.docx;*.xlsx;*.pptx"
      .FilterIndex = 1
      .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDocument = Nothing
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, _
URL As Variant)
   On Error Resume Next
   Set oDocument = pDisp.Document
   WebBrowser1.Document = pDisp.Document
End Sub

