Attribute VB_Name = "init"

'namespace=vba-files\Helper


Public Sub Alert()
    MsgBox "Alert.............", , "Title Alert"
End Sub



Public Sub AddNewFormData()

    Dim itemClient As String
    Dim itemDate As Date
    Dim itemAmount As Currency
    Dim itemClientType As String 

    
    itemClient = Sheets("Form").Range("Client").Value
    itemDate = Sheets("Form").Range("pDate").Value
    itemAmount = Sheets("Form").Range("Amount").Value
    itemClientType = Sheets("Form").Range("client_type").Value


    Sheets("Data").Activate 
    Range("A2").Select


    Do Until ActiveCell.Value = ""
        ActiveCell.Offset(1, 0).Select
    Loop


    ActiveCell.Value = itemClient
    ActiveCell.Offset(0, 1).Value = itemDate
    ActiveCell.Offset(0, 2).Value = itemAmount
    ActiveCell.Offset(0, 3).Value = itemClientType

    Sheet1.Range("alert_info").Value = "Data Submitted Successfully. " & Now()



End Sub



 