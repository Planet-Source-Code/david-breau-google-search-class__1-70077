VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGoogle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event result(oA As HTMLAnchorElement, sDescrip As String, _
             iresultnum As Integer, icurrpagenum As Integer)
Event docstate(sstate As String)
Event searchmatchednodocuments()
Event timeout()
Event done()
Event error(serr As String)

Const BASE_HREF = "http://www.google.com/search?hl=en&safe=off&num=100&q="
 

Private Type variables
  surl As String
  old_state As String
  odoc As HTMLDocument
  new_doc As New HTMLDocument
  itot_results_ret As Integer
  imax_results_ret As Integer
  icurr_pagenum As Integer
  imillisec_timeout As Integer
End Type
Dim v As variables

Private WithEvents mTmr As Timer
Attribute mTmr.VB_VarHelpID = -1

Public bCancel As Boolean


Function search(str_search_words As String, _
                 Optional inum_results As Integer = 100, _
                 Optional millisec_timeout As Integer = 10000, _
                 Optional yourtimer As Timer) As String
                 
On Error GoTo err_handler:

'google wont return more than 1000 results for any search
If inum_results < 1 Or inum_results > 1000 Then
  MsgBox "Must be a valid number from 1-1000", vbInformation
  Exit Function
End If

'make sure a search query was specified
If Len(Trim$(str_search_words)) = 0 Then
  MsgBox "Must provide search words to search for", vbInformation
  Exit Function
End If

If Not yourtimer Is Nothing Then
  v.imillisec_timeout = millisec_timeout
  Set mTmr = yourtimer
  mTmr.Interval = millisec_timeout
  mTmr.Enabled = True
End If

v.itot_results_ret = 0
v.imax_results_ret = inum_results
v.icurr_pagenum = 0
bCancel = False

'replace spaces with +
Dim s As String
s = Trim$(Replace$(str_search_words, " ", "+"))
v.surl = (BASE_HREF & s)
Call create_doc(v.surl)

Exit Function
err_handler:
  With Err
    If .Number <> 0 Then
      s = "cGoogle.search." & .Number & .Description
      RaiseEvent error(s)
    End If
End With
End Function


'---------------------------------------
'this function creates the document that
'is the google search results
'---------------------------------------
Private Function create_doc(surl As String) As String
On Error GoTo err_handler:
'
'create the google search results document
Dim new_doc As New HTMLDocument
Set v.odoc = v.new_doc.createDocumentFromUrl(surl, vbNullString)
'
'wait for it to be ready
With v.odoc
  While .readystate <> "complete"
    DoEvents
    If v.old_state <> .readystate Then
      v.old_state = .readystate
      RaiseEvent docstate(.readystate)
      If bCancel Then Exit Function
    End If
    DoEvents
  Wend
  
  If Not mTmr Is Nothing Then mTmr.Enabled = False
  RaiseEvent docstate(.readystate)
  
  '[bSearchMatchedNoDocuments] is a function below
  If bSearchMatchedNoDocuments Then
     RaiseEvent searchmatchednodocuments
     Set new_doc = Nothing
     Exit Function
  End If
  
  Set new_doc = Nothing
  Call parse_doc
End With

Exit Function
err_handler:
  With Err
    If .Number <> 0 Then
          Dim s As String
          s = "cGoogle.create_doc." & .Number & .Description
          RaiseEvent error(s)
    End If
  End With
End Function

'---------------------------------------
'this function parses the document examining
'the div elements. The div elements that
'have two children are most likely a search
'result so at that point we pass the div
'element onto the next function [extract_data]
'for further examination and then the
'passing the result info (the url and text)
'into the event raised [result]
'---------------------------------------
Private Function parse_doc() As String
On Error GoTo err_handler:
Dim icnt As Integer, i As Integer
 
  icnt = v.odoc.getElementsByTagName("div").length - 1
   
  If icnt = 0 Then
    RaiseEvent done
    Exit Function
  End If
  
 
  For i = 0 To icnt
     DoEvents
     
     If bCancel Then Exit Function
     
     'search results are contain within DIV in the doc
     Dim odiv As HTMLDivElement
     Set odiv = v.odoc.getElementsByTagName("div").Item(i)
     
     If Not odiv Is Nothing Then
       'how many childred nodes are there. the search
       'results contain two.
       If odiv.childNodes.length = 2 Then
          Call extract_data(odiv)
       End If
     End If
     
     DoEvents
  Next i
  
  
  If v.itot_results_ret >= v.imax_results_ret Then
     RaiseEvent done
  Else
     v.icurr_pagenum = (v.icurr_pagenum + 1)
     Debug.Print v.itot_results_ret
     Call create_doc(v.surl & "&start=" & CStr(v.icurr_pagenum) & "00")
  End If
  
Exit Function
err_handler:
  With Err
    If .Number <> 0 Then
          Dim s As String
          s = "cGoogle.parse_doc." & .Number & .Description
          RaiseEvent error(s)
    End If
  End With
End Function


Private Function extract_data(odiv As HTMLDivElement) As String
On Error GoTo err_handler:
Dim stag As String
  '
  'a google div element that is a search
  'return has two child nodes.
  'the first has a tagname of H2 and the second TABLE
   If LCase$(odiv.children(0).tagName) = "h2" Then
     If LCase$(odiv.children(1).tagName) = "table" Then
       '
       'keep track of total results returned so far
       If v.itot_results_ret >= v.imax_results_ret Then
         bCancel = True: DoEvents: RaiseEvent done
       Else
         v.itot_results_ret = (v.itot_results_ret + 1)
       End If
       
       RaiseEvent result(odiv.children(0).children(0), _
                  odiv.children(1).innerText, v.itot_results_ret, _
                  v.icurr_pagenum + 1)
                  
     End If
   End If
   
Exit Function
err_handler:
  With Err
    If .Number <> 0 Then
          Dim s As String
          s = "cGoogle.extract_data." & .Number & .Description
    End If
  End With
End Function

Private Function bSearchMatchedNoDocuments() As Boolean
On Error GoTo err_handler:
Dim stxt As String
stxt = LCase$(v.odoc.body.innerText)

If InStr(1, stxt, "your search") Then
  If InStr(1, stxt, "did not match any documents") Then
     bSearchMatchedNoDocuments = True
  End If
End If

Exit Function
err_handler:
  With Err
    If .Number <> 0 Then
          Dim s As String
          s = "cGoogle.bSearchMatchedNoDocuments." & .Number & "." & .Description
    End If
  End With
End Function

Private Sub Tmr_Timer()

bCancel = True
DoEvents
mTmr.Enabled = False
RaiseEvent timeout

End Sub

Private Sub Class_Terminate()

bCancel = True
DoEvents
DoEvents
Set v.odoc = Nothing
Set v.new_doc = Nothing

If Not mTmr Is Nothing Then
  mTmr.Enabled = False
  Set mTmr = Nothing
End If

End Sub
 
