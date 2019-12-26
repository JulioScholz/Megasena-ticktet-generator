Attribute VB_Name = "FormPost_Random_HTTP_API"
Option Explicit

Sub post_form()

    Dim http As Object
    Dim formURL As String, apiURL As String
    Dim id_header_name As String, id_key As String, api_resp As String, i As Range, num As Variant
    Dim Count As Integer, SplitCatcher As Variant, arr As Variant, str_range
    
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
    
    id_header_name = "Content-Type"
    id_key = "application/x-www-form-urlencoded; charset=utf-8"
    
    
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
                                                                                           
    apiURL = "https://www.random.org/sequences/?min=1&max=60&col=1&format=plain&rnd=new" & "&" & CInt(Rnd() * 10000)
    http.Open "GET", apiURL, False
    http.send
    
    'Random.org API response
    api_resp = http.responseText
    
    'Splitting the response
    SplitCatcher = Split(api_resp, vbLf)
    
    If num = 6 Then
        str_range = "A1:A6"
    ElseIf num = 7 Then
        str_range = "A1:A7"
    ElseIf num = 8 Then
        str_range = "A1:A8"
    ElseIf num = 9 Then
        str_range = "A1:A9"
    End If
    
    For Each i In Range("A1:A9")
      i.value = ""
      
    Next
    
    Count = 0
    For Each i In Range(str_range)
        i.value = SplitCatcher(Count)
        Count = Count + 1
    Next
    
   'The random numbers are sorted
    Columns("A").Sort key1:=Range("A1"), _
      order1:=xlAscending, Header:=xlNo
    
    arr = Array()
    ReDim Preserve arr(9)
    Count = 0
    For Each i In Range(str_range)
      arr(Count) = i.value
      Count = Count + 1
    Next
   
    
    ' The url is builded
    'Required entries must be modified with their google form entries
    formURL = "https://docs.google.com/forms/d/e/INSERT YOUT GOOGLE FORMS ID HERE?ifq"
    formURL = formURL & "&entry.1873463356=" & arr(0) & _
    "&entry.2088204043=" & arr(1) & _
    "&entry.1753095796=" & arr(2) & _
    "&entry.894483175=" & arr(3) & _
    "&entry.582438764=" & arr(4) & _
    "&entry.1154621635=" & arr(5) & _
    "&entry.1997354082=" & arr(6) & _
    "&entry.177761373=" & arr(7) & _
    "&entry.52510593=" & arr(8)
    
    'Posting data in the form
    http.Open "POST", formURL, False
    http.setRequestHeader id_header_name, id_key
    http.send


    'MsgBox http.responseText
        
    Set http = Nothing
    Exit Sub
End Sub


