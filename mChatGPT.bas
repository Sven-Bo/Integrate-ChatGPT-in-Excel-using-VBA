Option Explicit

'#################################################################################
'##  Title:   ChatGPT Completions using OpenAI API
'##  Author:  Sven from CodingIsFun
'##  Website: https://pythonandvba.com
'##  YouTube: https://youtube.com/@codingisfun
'##
'##  Description: This VBA script uses the OpenAI API endpoint "completions" to generate
'##               a text completion based on the selected cell and displays the result in a
'##               worksheet called OUTPUT_WORKSHEET. If the worksheet does not exist, it will be
'##               created. The API key is required to use the API, and it should be added as a
'##               constant at the top of the script.
'##               To get an API key, sign up for an OpenAI API key at https://openai.com/api/
'#################################################################################

'=====================================================
' GET YOUR API KEY: https://openai.com/api/
Const API_KEY As String = "<API_KEY>"
'=====================================================

' Constants for API endpoint and request properties
Const API_ENDPOINT As String = "https://api.openai.com/v1/completions"
Const MODEL As String = "text-davinci-003"
Const MAX_TOKENS As String = "1024"
Const TEMPERATURE As String = "0.5"

'Output worksheet name
Const OUTPUT_WORKSHEET As String = "Result"


Sub OpenAI_Completion()

    #If Mac Then
10            MsgBox "This macro only works on Windows. It is not compatible with macOS.", _
                  vbOKOnly, "Windows Compatibility Only"
20            Exit Sub
    #End If

30        On Error GoTo ErrorHandler
40        Application.screenupdating = False

          ' Check if API key is available
50        If API_KEY = "<API_KEY>" Then
60            MsgBox "Please input a valid API key. You can get one from https://openai.com/api/", vbCritical, "No API Key Found"
70            Application.screenupdating = True
80            Exit Sub
90        End If

          ' Get the prompt
          Dim prompt As String
          Dim cell As Range
          Dim selectedRange As Range
100       Set selectedRange = Selection
          
110       For Each cell In selectedRange
120           prompt = prompt & cell.Value & " "
130       Next cell

          ' Check if there is anything in the selected cell
140       If Trim(prompt) <> "" Then
              ' Clean prompt to avoid parsing error in JSON payload
150           prompt = CleanJSONString(prompt)
160       Else
170           MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
180           Application.screenupdating = True
190           Exit Sub
200       End If

          ' Create worksheet if it does not exist
210       If Not WorksheetExists(OUTPUT_WORKSHEET) Then
220           Worksheets.Add(After:=Sheets(Sheets.Count)).name = OUTPUT_WORKSHEET
230       End If

          ' Clear existing data in worksheet
240       Worksheets(OUTPUT_WORKSHEET).UsedRange.ClearContents

          ' Show status in status bar
250       Application.StatusBar = "Processing OpenAI request..."

          ' Create XMLHTTP object
          Dim httpRequest As Object
260       Set httpRequest = CreateObject("MSXML2.XMLHTTP")

          ' Define request body
          Dim requestBody As String
270       requestBody = "{" & _
              """model"": """ & MODEL & """," & _
              """prompt"": """ & prompt & """," & _
              """max_tokens"": " & MAX_TOKENS & "," & _
              """temperature"": " & TEMPERATURE & _
              "}"
              
          ' Open and send the HTTP request
280       With httpRequest
290           .Open "POST", API_ENDPOINT, False
300           .SetRequestHeader "Content-Type", "application/json"
310           .SetRequestHeader "Authorization", "Bearer " & API_KEY
320           .send (requestBody)
330       End With

          'Check if the request is successful
340       If httpRequest.Status = 200 Then
              'Parse the JSON response
              Dim response As String
350           response = httpRequest.responseText

              'Get the completion and clean it up
              Dim completion As String
360           completion = ParseResponse(response)
              
              'Split the completion into lines
              Dim lines As Variant
370           lines = Split(completion, "\n")

              'Write the lines to the worksheet
              Dim i As Long
380           For i = LBound(lines) To UBound(lines)
390               Worksheets(OUTPUT_WORKSHEET).Cells(i + 1, 1).Value = ReplaceBackslash(lines(i))
400           Next i

              'Auto fit the column width
410           Worksheets(OUTPUT_WORKSHEET).Columns.AutoFit
              
              ' Show completion message
420           MsgBox "OpenAI completion request processed successfully. Results can be found in the 'Result' worksheet.", vbInformation, "OpenAI Request Completed"
              
              'Activate & color result worksheet
430           With Worksheets(OUTPUT_WORKSHEET)
440               .Activate
450               .Range("A1").Select
460               .Tab.Color = RGB(169, 208, 142)
470           End With
              
480       Else
490           MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
500       End If
          
510       Application.StatusBar = False
520       Application.screenupdating = True
          
530       Exit Sub
          
ErrorHandler:
540       MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical, "Error"
550       Application.StatusBar = False
560       Application.screenupdating = True
          
End Sub
' Helper function to check if worksheet exists
Function WorksheetExists(worksheetName As String) As Boolean
570       On Error Resume Next
580       WorksheetExists = (Not (Sheets(worksheetName) Is Nothing))
590       On Error GoTo 0
End Function
' Helper function to parse the reponse text
Function ParseResponse(ByVal response As String) As String
600       On Error Resume Next
          Dim startIndex As Long
610       startIndex = InStr(response, """text"":""") + 8
          Dim endIndex As Long
620       endIndex = InStr(response, """index"":") - 2
630       ParseResponse = Mid(response, startIndex, endIndex - startIndex)
640       On Error GoTo 0
End Function
' Helper function to clean text
Function CleanJSONString(inputStr As String) As String
650       On Error Resume Next
          ' Remove line breaks
660       CleanJSONString = Replace(inputStr, vbCrLf, "")
670       CleanJSONString = Replace(CleanJSONString, vbCr, "")
680       CleanJSONString = Replace(CleanJSONString, vbLf, "")

          ' Replace all double quotes with single quotes
690       CleanJSONString = Replace(CleanJSONString, """", "'")
700       On Error GoTo 0
End Function
' Replaces the backslash character only if it is immediately followed by a double quote.
Function ReplaceBackslash(text As Variant) As String
710       On Error Resume Next
          Dim i As Integer
          Dim newText As String
720       newText = ""
730       For i = 1 To Len(text)
740           If Mid(text, i, 2) = "\" & Chr(34) Then
750               newText = newText & Chr(34)
760               i = i + 1
770           Else
780               newText = newText & Mid(text, i, 1)
790           End If
800       Next i
810       ReplaceBackslash = newText
820       On Error GoTo 0
End Function
