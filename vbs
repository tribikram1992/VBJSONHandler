Function ParseJson(jsonStr)
    Dim stack: Set stack = CreateObject("Scripting.Dictionary")
    Dim result: Set result = CreateObject("Scripting.Dictionary")
    Dim currentKey
    Dim currentValue
    Dim inString
    Dim escape
    Dim i

    i = 1
    While i <= Len(jsonStr)
        Dim char: char = Mid(jsonStr, i, 1)

        If inString Then
            If escape Then
                escape = False
            ElseIf char = "\" Then
                escape = True
            ElseIf char = """" Then
                inString = False
            End If
            currentValue = currentValue & char
        Else
            Select Case char
                Case "{"
                    If Not IsEmpty(currentKey) Then
                        stack.Add stack.Count, result
                        result.Add currentKey, CreateObject("Scripting.Dictionary")
                        Set result = result(currentKey)
                        currentKey = Empty
                    Else
                        stack.Add stack.Count, CreateObject("Scripting.Dictionary")
                        Set result = stack(stack.Count - 1)
                    End If
                Case "}"
                    If Not IsEmpty(currentKey) Then
                        result.Add currentKey, currentValue
                        currentKey = Empty
                        currentValue = Empty
                    End If
                    If stack.Count > 0 Then
                        Set result = stack(stack.Count - 1)
                        stack.Remove stack.Count - 1
                    End If
                Case "["
                    If Not IsEmpty(currentKey) Then
                        stack.Add stack.Count, result
                        result.Add currentKey, CreateObject("Scripting.Dictionary")
                        Set result = result(currentKey)
                        currentKey = Empty
                    Else
                        stack.Add stack.Count, CreateObject("Scripting.Dictionary")
                        Set result = stack(stack.Count - 1)
                    End If
                Case "]"
                    If Not IsEmpty(currentValue) Then
                        result.Add result.Count, currentValue
                        currentValue = Empty
                    End If
                    If stack.Count > 0 Then
                        Set result = stack(stack.Count - 1)
                        stack.Remove stack.Count - 1
                    End If
                Case ":"
                    currentKey = currentValue
                    currentValue = Empty
                Case ","
                    If Not IsEmpty(currentValue) Then
                        If IsEmpty(currentKey) Then
                            result.Add result.Count, currentValue
                        Else
                            result.Add currentKey, currentValue
                        End If
                        currentValue = Empty
                    End If
                Case """"
                    inString = True
                    currentValue = ""
                Case " ", vbTab, vbCr, vbLf
                    ' Ignore whitespace
                Case Else
                    currentValue = currentValue & char
            End Select
        End If

        i = i + 1
    Wend

    ' Final check for any remaining value
    If Not IsEmpty(currentKey) And Not IsEmpty(currentValue) Then
        result.Add currentKey, currentValue
    End If

    Set ParseJson = result
End Function

' Example usage:
Dim jsonString
jsonString = "{""name"": ""John Smith"", ""age"": 30, ""city"": ""New York"", ""children"": [{""name"": ""Jane"", ""age"": 5}, {""name"": ""Alex"", ""age"": 8}]}"
Dim jsonObject
Set jsonObject = ParseJson(jsonString)

' Output the parsed JSON object
Dim key
For Each key In jsonObject.Keys
    WScript.Echo key & ": " & jsonObject(key)
Next
