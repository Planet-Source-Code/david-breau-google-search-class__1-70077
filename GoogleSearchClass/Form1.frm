VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstUrls 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   5265
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Left            =   2655
      Top             =   2070
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "dbl click to launch result in your browser"
      Top             =   585
      Width           =   7080
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "free vb controls"
      Top             =   225
      Width           =   2625
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search Google"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   225
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private WithEvents cGoogle As cGoogle
Attribute cGoogle.VB_VarHelpID = -1


Private Sub Form_Load()
  Set cGoogle = New cGoogle
End Sub


Private Sub cGoogle_docstate(sstate As String)
  Debug.Print sstate
End Sub

Private Sub cGoogle_done()
  'we have extracted the number of result desired
  'as set by [inum_results] of  cgoogle.search
  MsgBox "Done!!" & vbCrLf & _
          List1.ListCount & " results returned"
End Sub

Private Sub cGoogle_error(serr As String)
  Beep
  Debug.Print serr
End Sub

Private Sub cGoogle_result(oA As MSHTML.HTMLAnchorElement, _
                      sDescrip As String, iresultnum As Integer, _
                          icurrpagenum As Integer)
   'a search result returned
   List1.AddItem Format(iresultnum, "000") & "   " & sDescrip
   lstUrls.AddItem oA.href
End Sub

Private Sub cGoogle_searchmatchednodocuments()
  MsgBox "No Results Returned For Your Search"
End Sub

Private Sub cGoogle_timeout()
  MsgBox "Timed Out"
End Sub

Private Sub cmdSearch_Click()
  List1.Clear
  lstUrls.Clear
  cGoogle.search txtSearch, 275, , Timer1
End Sub
 

Private Sub List1_DblClick()
  With lstUrls
     If List1.ListIndex >= 0 Then
        ShellExecute hwnd, "open", .List(List1.ListIndex), "", "", 1
     End If
  End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set cGoogle = Nothing
End Sub
