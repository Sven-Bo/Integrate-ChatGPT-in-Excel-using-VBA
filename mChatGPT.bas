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
Const MAX_TOKENS As Long = 1024
Const TEMPERATURE As Double = 0.5

'Output worksheet name
Const OUTPUT_WORKSHEET As String = "Result"


Sub OpenAI_Completion()

10        On Error GoTo ErrorHandler
20        Application.ScreenUpdating = False

          ' Check if API key is available
30        If API_KEY = "<API_KEY>" Then
40            MsgBox "Please input a valid API key. You can get one from https://openai.com/api/", vbCritical, "No API Key Found"
50            Application.ScreenUpdating = False
60            Exit Sub
70        End If

          ' Get the prompt
          Dim prompt As String
80        prompt = ActiveCell.Value

          ' Check if there is anything in the selected cell
90        If Trim(prompt) <> "" Then
              ' Clean prompt to avoid parsing error in JSON payload
100           prompt = CleanJSONString(prompt)
110       Else
120           MsgBox "Please enter some text in the selected cell before executing the macro", vbCritical, "Empty Input"
130           Application.ScreenUpdating = False
140           Exit Sub
150       End If

          ' Create worksheet if it does not exist
160       If Not WorksheetExists(OUTPUT_WORKSHEET) Then
170           Worksheets.Add(After:=Sheets(Sheets.Count)).Name = OUTPUT_WORKSHEET
180       End If

          ' Clear existing data in worksheet
190       Worksheets(OUTPUT_WORKSHEET).UsedRange.ClearContents

          ' Show status in status bar
200       Application.StatusBar = "Processing OpenAI request..."

          ' Create XMLHTTP object
          Dim httpRequest As Object
210       Set httpRequest = CreateObject("MSXML2.XMLHTTP")

          ' Define request body
          Dim requestBody As String
220       requestBody = "{" & _
              """model"": """ & MODEL & """," & _
              """prompt"": """ & prompt & """," & _
              """max_tokens"": " & MAX_TOKENS & "," & _
              """temperature"": " & TEMPERATURE & _
              "}"
              
          ' Open and send the HTTP request
230       With httpRequest
240           .Open "POST", API_ENDPOINT, False
250           .SetRequestHeader "Content-Type", "application/json"
260           .SetRequestHeader "Authorization", "Bearer " & API_KEY
270           .send (requestBody)
280       End With

          'Check if the request is successful
290       If httpRequest.Status = 200 Then
              'Parse the JSON response
              Dim response As String
300           response = httpRequest.responseText

              'Get the completion and clean it up
              Dim completion As String
310           completion = ParseResponse(response)
              
              'Split the completion into lines
              Dim lines As Variant
320           lines = Split(completion, "\n")

              'Write the lines to the worksheet
              Dim i As Long
330           For i = LBound(lines) To UBound(lines)
340               Worksheets(OUTPUT_WORKSHEET).Cells(i + 1, 1).Value = ReplaceBackslash(lines(i))
350           Next i

              'Auto fit the column width
360           Worksheets(OUTPUT_WORKSHEET).Columns.AutoFit
              
              ' Show completion message
370           MsgBox "OpenAI completion request processed successfully. Results can be found in the 'Result' worksheet.", vbInformation, "OpenAI Request Completed"
              
              'Activate & color result worksheet
380           With Worksheets(OUTPUT_WORKSHEET)
390               .Activate
400               .Range("A1").Select
410               .Tab.Color = RGB(169, 208, 142)
420           End With
              
430       Else
440           MsgBox "Request failed with status " & httpRequest.Status & vbCrLf & vbCrLf & "ERROR MESSAGE:" & vbCrLf & httpRequest.responseText, vbCritical, "OpenAI Request Failed"
450       End If
          
460       Application.StatusBar = False
470       Application.ScreenUpdating = True
          
480       Exit Sub
          
ErrorHandler:
490       MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical, "Error"
500       Application.StatusBar = False
510       Application.ScreenUpdating = True
          
End Sub
' Helper function to check if worksheet exists
Function WorksheetExists(worksheetName As String) As Boolean
520       On Error Resume Next
530       WorksheetExists = (Not (Sheets(worksheetName) Is Nothing))
540       On Error GoTo 0
End Function
' Helper function to parse the reponse text
Function ParseResponse(ByVal response As String) As String
550       On Error Resume Next
          Dim startIndex As Long
560       startIndex = InStr(response, """text"":""") + 8
          Dim endIndex As Long
570       endIndex = InStr(response, """index"":") - 2
580       ParseResponse = Mid(response, startIndex, endIndex - startIndex)
590       On Error GoTo 0
End Function
' Helper function to clean text
Function CleanJSONString(inputStr As String) As String
600       On Error Resume Next
          ' Remove line breaks
610       CleanJSONString = Replace(inputStr, vbCrLf, "")
620       CleanJSONString = Replace(CleanJSONString, vbCr, "")
630       CleanJSONString = Replace(CleanJSONString, vbLf, "")

          ' Replace all double quotes with single quotes
640       CleanJSONString = Replace(CleanJSONString, """", "'")
650       On Error GoTo 0
End Function
' Replaces the backslash character only if it is immediately followed by a double quote.
Function ReplaceBackslash(text As Variant) As String
660       On Error Resume Next
          Dim i As Integer
          Dim newText As String
670       newText = ""
680       For i = 1 To Len(text)
690           If Mid(text, i, 2) = "\" & Chr(34) Then
700               newText = newText & Chr(34)
710               i = i + 1
720           Else
730               newText = newText & Mid(text, i, 1)
740           End If
750       Next i
760       ReplaceBackslash = newText
770       On Error GoTo 0
End Function


