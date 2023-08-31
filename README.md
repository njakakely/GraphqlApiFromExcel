# GraphqlApiFromExcel

This code is used to fetch GraphQLApi from excel

## The code on Vba

Sub querygraph()

Dim url As String
Dim query As String
Dim request As Object
Dim response As String

url = "http://localhost:4000/api"

query = '{ users { _id nom prenom }}'

Set request = CreateObject("MSXML2.XMLHTTP.6.0")<br>
request.Open "POST", url, False<br>
request.setrequestheader "Content-Type", "application/json"<br>
request.send "{ ""query"": """ & query & """ }"<br>

response = request.responseText<br>

MsgBox response<br>


End Sub
