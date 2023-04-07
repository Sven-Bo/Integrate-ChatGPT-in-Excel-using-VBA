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

10        On Error GoTo ErrorHandler
20        Application.ScreenUpdating = False

          ' Check if API key is available
30        If API_KEY = "<API_KEY>" Then
40            MsgBox "Please input a valid API key. You can get one from https://openai.com/api/", vbCritical, "No API Key Found"
50            Application.ScreenUpdating = True
60            Exit Sub
70        End If

          ' Get the prompt
          Dim prompt As String
          Dim cell As Range
          Dim selectedRange As Range
80        Set selectedRange = Selection
          
90        For Each cell In selectedRange
100           prompt = prompt & cell.Value & " "
110       Next cell

          ' Check if there is anything in the selected cell
120       If Trim(prompt) <> "" Then
              ' Clean prompt to avoid parsing error in JSON payload
130           prompt = CleanJSONString(prompt)
140       Else
150           MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
160           Application.ScreenUpdating = True
170           Exit Sub
180       End If

          ' Create worksheet if it does not exist
190       If Not WorksheetExists(OUTPUT_WORKSHEET) Then
200           Worksheets.Add(After:=Sheets(Sheets.Count)).Name = OUTPUT_WORKSHEET
210       End If

          ' Clear existing data in worksheet
220       Worksheets(OUTPUT_WORKSHEET).UsedRange.ClearContents

          ' Show status in status bar
230       Application.StatusBar = "Processing OpenAI request..."

          ' Create XMLHTTP object
          Dim httpRequest As Object
240       Set httpRequest = CreateObject("MSXML2.XMLHTTP")

          ' Define request body
          Dim requestBody As String
250       requestBody = "{" & _
              """model"": """ & MODEL & """," & _
              """prompt"": """ & prompt & """," & _
              """max_tokens"": " & MAX_TOKENS & "," & _
              """temperature"": " & TEMPERATURE & _
              "}"
              
          ' Open and send the HTTP request
260       With httpRequest
270           .Open "POST", API_ENDPOINT, False
280           .SetRequestHeader "Content-Type", "application/json"
290           .SetRequestHeader "Authorization", "Bearer " & API_KEY
300           .send (requestBody)
310       End With

          'Check if the request is successful
320       If httpRequest.Status = 200 Then
              'Parse the JSON response
              Dim response As String
330           response = httpRequest.responseText

              'Get the completion and clean it up
              Dim completion As String
340           completion = ParseResponse(response)
              
              'Split the completion into lines
              Dim lines As Variant
350           lines = Split(completion, "\n")

              'Write the lines to the worksheet
              Dim i As Long
360           For i = LBound(lines) To UBound(lines)
370               Worksheets(OUTPUT_WORKSHEET).Cells(i + 1, 1).Value = ReplaceBackslash(lines(i))
380           Next i

              'Auto fit the column width
390           Worksheets(OUTPUT_WORKSHEET).Columns.AutoFit
              
              ' Show completion message
400           MsgBox "OpenAI completion request processed successfully. Results can be found in the 'Result' worksheet.", vbInformation, "OpenAI Request Completed"
              
              'Activate & color result worksheet
410           With Worksheets(OUTPUT_WORKSHEET)
420               .Activate
430               .Range("A1").Select
440               .Tab.Color = RGB(169, 208, 142)
450           End With
              
460       Else
470           MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
480       End If
          
490       Application.StatusBar = False
500       Application.ScreenUpdating = True
          
510       Exit Sub
          
ErrorHandler:
520       MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical, "Error"
530       Application.StatusBar = False
540       Application.ScreenUpdating = True
          
End Sub
' Helper function to check if worksheet exists
Function WorksheetExists(worksheetName As String) As Boolean
550       On Error Resume Next
560       WorksheetExists = (Not (Sheets(worksheetName) Is Nothing))
570       On Error GoTo 0
End Function
' Helper function to parse the reponse text
Function ParseResponse(ByVal response As String) As String
580       On Error Resume Next
          Dim startIndex As Long
590       startIndex = InStr(response, """text"":""") + 8
          Dim endIndex As Long
600       endIndex = InStr(response, """index"":") - 2
610       ParseResponse = Mid(response, startIndex, endIndex - startIndex)
620       On Error GoTo 0
End Function
' Helper function to clean text
Function CleanJSONString(inputStr As String) As String
630       On Error Resume Next
          ' Remove line breaks
640       CleanJSONString = Replace(inputStr, vbCrLf, "")
650       CleanJSONString = Replace(CleanJSONString, vbCr, "")
660       CleanJSONString = Replace(CleanJSONString, vbLf, "")

          ' Replace all double quotes with single quotes
670       CleanJSONString = Replace(CleanJSONString, """", "'")
680       On Error GoTo 0
End Function
' Replaces the backslash character only if it is immediately followed by a double quote.
Function ReplaceBackslash(text As Variant) As String
690       On Error Resume Next
          Dim i As Integer
          Dim newText As String
700       newText = ""
710       For i = 1 To Len(text)
720           If Mid(text, i, 2) = "\" & Chr(34) Then
730               newText = newText & Chr(34)
740               i = i + 1
750           Else
760               newText = newText & Mid(text, i, 1)
770           End If
780       Next i
790       ReplaceBackslash = newText
800       On Error GoTo 0
End Function
