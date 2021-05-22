Attribute VB_Name = "Parrot"
Option Explicit

Function ParrotBINGet( _
        ByVal sUrl As String, _
        ByVal sApiKey As String, _
        ByVal sBIN As String, _
        Optional ByVal paths As String = "/card/number/length," & _
            "/card/number/luhn," & _
            "/card/scheme," & _
            "/card/type," & _
            "/card/brand," & _
            "/card/prepaid," & _
            "/card/country/numeric," & _
            "/card/country/alpha2," & _
            "/card/country/name," & _
            "/card/country/emoji," & _
            "/card/country/currency," & _
            "/card/country/latitude," & _
            "/card/country/longitude," & _
            "/card/bank/name," & _
            "/card/bank/url," & _
            "/card/bank/phone," & _
            "/card/bank/city", _
            Optional ByVal useCache As Boolean = True) As String()
    
    
    Static result_cache As New Dictionary
    
    
    Dim Client As New WebClient
    Client.BaseUrl = sUrl
    Dim Request As New WebRequest
    Request.Method = WebMethod.HTTPGet
    Request.ResponseFormat = WebFormat.Json
    Request.AddHeader "x-api-key", sApiKey
    Request.AddQuerystringParam "bin", sBIN
    Request.AddQuerystringParam "enhanced", "false"
    
    Dim Response As WebResponse
    
    Dim jsonText As String
    Dim pathArr() As String
    Dim JsonResponse As Object
    Dim results() As String
    
    
    Dim keycheck As String
    keycheck = sBIN
    
    If useCache Then
        jsonText = result_cache(keycheck)
    Else
        jsonText = ""
    End If
    
    If jsonText = "" Then
        'Debug.print sUrl, sBIN
        Set Response = Client.Execute(Request)
        jsonText = Response.Content
        
        If (result_cache.Exists(keycheck)) Then
            result_cache.Remove (keycheck)
        End If
        
        result_cache.Add keycheck, jsonText
        'Debug.print jsonText
    Else
        ''Debug.print "CACHED", jsonText
    End If
    Set JsonResponse = ParseJson(jsonText)
    
    ''Debug.print JsonResponse("card")("length")
    
    
    
   Dim errortext As String
   
BINNOTFOUND:
   On Error GoTo TOOMANY
   errortext = JsonResponse("result")("info")
TOOMANY:
    On Error GoTo -1
    
    On Error GoTo CONTINUE
    If (errortext = "") Then
        errortext = JsonResponse("message")
    End If
    
    
    
    If (errortext = "Bin Card does not exist in our database!") Then
             ' don't cache not found
             'If (result_cache.Exists(keycheck)) Then
             '    result_cache.Remove (keycheck)
             'End If
             ReDim results(1)
             results(0) = "Sorry, BIN is not found."
             ParrotBINGet = results
             Exit Function
    ElseIf (errortext = "BIN card number is not valid.") Then
             ReDim results(1)
             results(0) = "Not valid"
             ParrotBINGet = results
             Exit Function
   ElseIf (errortext = "Requested object was not found : card was not found.") Then
             If (result_cache.Exists(keycheck)) Then
                 result_cache.Remove (keycheck)
             End If
             ReDim results(1)
             results(0) = "Sorry, BIN is not found."
             ParrotBINGet = results
             Exit Function
    ElseIf (errortext = "Unauthorized") Then
             ReDim results(1)
             results(0) = "Not authorized"
             ParrotBINGet = results
             Exit Function
   ElseIf (errortext = "Trial period has expired.") Then
             ReDim results(1)
             results(0) = "Free Trial Expired"
             ParrotBINGet = results
             Exit Function
    ElseIf (errortext = "Too many requests") Then
                 ' don't cache too many requests
             'If (result_cache.Exists(keycheck)) Then
             '    result_cache.Remove (keycheck)
             'End If
             ReDim results(1)
             results(0) = "Too many requests"
             ParrotBINGet = results
             Exit Function
    End If
    
CONTINUE:
    '''Debug.print "Path String: " & Len(paths); ""
    If (Len(paths) > 0) Then
        pathArr = Split(paths, ",")
    End If
   
    'Debug.print "Number of paths " & UBound(pathArr)
   
    If UBound(pathArr) = 16 Then
        ParrotBINGet = results
    Else
        ''Debug.print (paths)
        ReDim results(UBound(pathArr))
    End If
   
    Dim pt As Variant
    Dim cNode As Variant
    Dim pac As Integer

On Error GoTo LOOPERROR

    For pac = 0 To UBound(pathArr)
        pt = pathArr(pac)
        Set cNode = JsonResponse
        Dim jpath() As String
        jpath = Split(Mid(pt, 2), "/")
        Dim jp As Variant
        Dim pc As Variant
        'Debug.print "Working on Path: " & pt & " Size: " & UBound(jpath)
   
        For pc = 0 To UBound(jpath)
            Dim kval As Variant
            kval = jpath(pc)
            'Debug.print kval, pc
            
            If pc = UBound(jpath) Then
                If cNode.Exists(kval) And Not IsNull(cNode(kval)) Then
                    'Debug.print "Adding", cNode(kval)
                    results(pac) = cNode(kval)
                End If
            Else
                If Not IsEmpty(cNode(kval)) Then
                    Set cNode = cNode(kval)
                End If
            End If
        Next
   Next
       
       
    ParrotBINGet = results
    Exit Function
    
LOOPERROR:
  MsgBox "There seems to be an error" & vbCrLf & Err.Description
    
End Function