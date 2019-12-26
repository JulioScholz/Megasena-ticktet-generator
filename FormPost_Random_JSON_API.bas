Attribute VB_Name = "FormPost_Random_JSON_API"
Option Explicit

Sub post_form_json()

    Dim http As Object
    Dim formURL As String, apiURL As String
    Dim i As Range, num As Variant
    Dim Count As Integer, arr As Variant, str_range As String
    Dim Json As Dictionary, json_String As String
    
    '------------------------------------------------------------------------------------
    num = InputBox("Insira o número de dezenas que você deseja (máximo 9, mínimo 6) ", _
                Title:="No. de dezenas", Default:="Insira o número aqui")

    If num = "" Then
        Exit Sub
    ElseIf Not IsNumeric(num) Then
        MsgBox "You must enter a numerical value."
        Exit Sub
 
    ElseIf num > 9 Then
        MsgBox "Apenas 9 dezenas são permitidas"
        Exit Sub
    
    ElseIf num < 6 Then
        MsgBox "No mínimo 6 dezenas são permitidas"
        Exit Sub
    End If
    '------------------------------------------------------------------------------------
    
    json_String = " {""jsonrpc"": ""2.0"", " & _
    " ""method"": ""generateIntegerSequences""," & _
    " ""params"": { ""apiKey"": ""INSERT YOUR RANDOM.ORG API KEY HERE"", " & _
    " ""n"": 1 ,""length"": " & num & " ," & _
    " ""min"": 1,""max"": 60,""replacement"": false,""base"": 10 },""id"": 1}"

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    apiURL = "https://api.random.org/json-rpc/2/invoke"
    With http
        .Open "POST", apiURL, False
        .setRequestHeader "Content-Type", "application/json"
        'Sending the json in string format
        .send (json_String)
    End With
    
    'Random.org API response is converted to JSON
    Set Json = JsonConverter.ParseJson(http.responseText)
  
    'The range is setted based in the value of num
    str_range = "A1:A" + num
    For Each i In Range("A1:A9")
        i.value = ""
    Next
    
    'The sheet is populated
    Count = 1
    For Each i In Range(str_range)
        i.value = Json("result")("random")("data")(1)(Count)
        Count = Count + 1
    Next
    
   'The random numbers are sorted in the sheet
    Columns("A").Sort key1:=Range("A1"), _
      order1:=xlAscending, Header:=xlNo
    
    'The array is populated
    arr = Array()
    ReDim Preserve arr(9)
    Count = 0
    For Each i In Range(str_range)
        arr(Count) = i.value
        Count = Count + 1
    Next
   
    'The url is builded
    'Required entries must be modified with their google form entries
    formURL = "https://docs.google.com/forms/d/e/INSERT YOUT GOOGLE FORMS ID HERE?ifq"
    formURL = formURL & _
        "&entry.1873463356=" & arr(0) & _
        "&entry.2088204043=" & arr(1) & _
        "&entry.1753095796=" & arr(2) & _
        "&entry.894483175=" & arr(3) & _
        "&entry.582438764=" & arr(4) & _
        "&entry.1154621635=" & arr(5) & _
        "&entry.1997354082=" & arr(6) & _
        "&entry.177761373=" & arr(7) & _
        "&entry.52510593=" & arr(8)
        
    'Posting data in the form
    With http
        .Open "POST", formURL, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"
        .send
    End With
    Set http = Nothing
    
End Sub
