VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Function MainLoop()

'Get Bottom of Range and set Objects and Variables
lastRow = Cells(Rows.Count, "A").End(xlUp).Row
myRange = "A2:A" & lastRow
urlFirstPart = "https://www.cryptocurrencychart.com/coin-row/"
Set FinHTTP = CreateObject("Microsoft.xmlHTTP")



'Begin looping through our Tickers
For Counter = 2 To lastRow
    Ticker = Worksheets("Sheet1").Cells(Counter, 1).Value
    
    'Get The Price Data from the web
    UrlString = urlFirstPart & Ticker
    FinHTTP.Open "GET", UrlString, False
    FinHTTP.send
    htmlData = FinHTTP.responseText

    'Submit the raw html to a helper function which cleans it up and gives us the price
    coinPrice = TextCleaner(htmlData)
    
    'Set the price and market value cells
    Worksheets("Sheet1").Cells(Counter, 6).Value = coinPrice
    Worksheets("Sheet1").Cells(Counter, 7).Formula = "=E" & Counter & "*F" & Counter
    
Next Counter

'Sum up the data into a total value
Worksheets("Sheet1").Cells(lastRow + 1, 7).Formula = "=SUM(G2:G" & lastRow & ")"
Worksheets("Sheet1").Cells(lastRow + 1, 7).Font.Bold = True

End Function


Function TextCleaner(htmlResponseText)

    StartNum = InStr(1, htmlResponseText, "numeric price", vbTextCompare)
    EndNum = InStr(StartNum, htmlResponseText, "title", vbTextCompare)
    TextCleaner = Mid(htmlResponseText, StartNum + 29, EndNum - StartNum - 32)
    
End Function

Private Sub CommandButton1_Click()

MainLoop

End Sub
