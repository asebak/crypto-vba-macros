Attribute VB_Name = "Module"
Public Sub BTCPrice()

Dim oJSON As Object
Dim httpObject As Object
Set httpObject = CreateObject("MSXML2.XMLHTTP")
Dim rep As Variant
Dim var As Object

sURL = "https://api.coinbase.com/v2/prices/BTC-USD/spot"

sRequest = sURL
httpObject.Open "GET", sRequest, False
httpObject.Send
sGetResult = httpObject.responsetext
Dim Json As Object
Set Json = JsonConverter.ParseJson(sGetResult)
Set var = Json("data")
Debug.Print var("amount")
Dim ws As Worksheet
Set ws = Sheets("Sheet1")
ws.Range("D5").Value = var("amount")

End Sub

Public Sub ETHPrice()

Dim oJSON As Object
Dim httpObject As Object
Set httpObject = CreateObject("MSXML2.XMLHTTP")
Dim rep As Variant
Dim var As Object

sURL = "https://api.coinbase.com/v2/prices/ETH-USD/spot"

sRequest = sURL
httpObject.Open "GET", sRequest, False
httpObject.Send
sGetResult = httpObject.responsetext
Dim Json As Object
Set Json = JsonConverter.ParseJson(sGetResult)
Set var = Json("data")
Debug.Print var("amount")
Dim ws As Worksheet
Set ws = Sheets("Sheet1")
ws.Range("G5").Value = var("amount")

End Sub


