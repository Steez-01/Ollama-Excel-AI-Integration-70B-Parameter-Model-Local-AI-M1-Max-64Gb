# Ollama-Excel-AI-Integration-70B-Parameter-Model-Local-AI-M1-Max-64Gb
Ollama Excel AI Integration 70B Parameter Model Local AI M1 Max 64Gb causes you to save your AI Credits ! Since you let Local AI do the calculations !! Optimizing your pre-investment in your Laptop

70B Parameter Model Overview Code 

Function LocalAI(prompt As String) As String
    Dim cleanPrompt As String
    Dim shellCmd As String
    Dim response As String
    
    ' 1. Remove any line breaks and double quotes that break the JSON
    cleanPrompt = Replace(prompt, """", "'")
    cleanPrompt = Replace(cleanPrompt, vbLf, " ")
    cleanPrompt = Replace(cleanPrompt, vbCr, " ")
    
    ' 2. Build the curl command with STRICT JSON formatting
    ' We use ' around the whole data block to protect it from the shell
    shellCmd = "curl -s -X POST http://127.0.0.1:11434/api/generate " & _
               "-d ""{\""model\"": \""llama3.3\"", \""prompt\"": \""" & cleanPrompt & "\"", \""stream\"": false}"""
    
    ' 3. Call your working script
    On Error Resume Next
    response = AppleScriptTask("RunShell.scpt", "executeOllama", shellCmd)
    
    ' 4. Extract the response
    If InStr(response, """response"":""") > 0 Then
        LocalAI = Split(Split(response, """response"":""")(1), """")(0)
    Else
        LocalAI = "Terminal Error: " & response
    End If
End Function


8B Parameter Model Overview Code

Function LocalAI(prompt As String) As String
    Dim cleanPrompt As String
    Dim shellCmd As String
    Dim response As String
    
    ' 1. Remove any line breaks and double quotes that break the JSON
    cleanPrompt = Replace(prompt, """", "'")
    cleanPrompt = Replace(cleanPrompt, vbLf, " ")
    cleanPrompt = Replace(cleanPrompt, vbCr, " ")
    
    ' 2. Build the curl command with STRICT JSON formatting
    ' We use ' around the whole data block to protect it from the shell
    shellCmd = "curl -s -X POST http://127.0.0.1:11434/api/generate " & _
               "-d ""{\""model\"": \""llama3.3\"", \""prompt\"": \""" & cleanPrompt & "\"", \""stream\"": false}"""
    
    ' 3. Call your working script
    On Error Resume Next
    response = AppleScriptTask("RunShell.scpt", "executeOllama", shellCmd)
    
    ' 4. Extract the response
    If InStr(response, """response"":""") > 0 Then
        LocalAI = Split(Split(response, """response"":""")(1), """")(0)
    Else
        LocalAI = "Terminal Error: " & response
    End If
End Function
