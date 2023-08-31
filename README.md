# GraphqlApiFromExcel

This code is used to fetch GraphQLApi from excel

Sub querygraph()

Dim url As String
Dim query As String
Dim request As Object
Dim response As String

url = "http://localhost:4000/api"

query = '{ users { _id nom prenom }}'

Set request = CreateObject("MSXML2.XMLHTTP.6.0")
request.Open "POST", url, False
request.setrequestheader "Content-Type", "application/json"
request.send "{ ""query"": """ & query & """ }"

response = request.responseText

MsgBox response


End Sub
