Attribute VB_Name = "init"

'namespace=vba-files\Helper


Public Sub Alert()
    MsgBox "Alert.............", , "Title Alert"
End Sub



Public Sub AddNewData()
    Sheets("Data").Activate 
    Range("A2").Select

    Do Until ActiveCell.Value = ""
        ActiveCell.Offset(1, 0).Select
    Loop

    ActiveCell.Value = Sheets("Form").Range("Client").Value
    ActiveCell.Offset(0, 1).Value = Sheets("Form").Range("pDate").Value
    ActiveCell.Offset(0, 2).Value = Sheets("Form").Range("Amount").Value

    Sheet1.Range("C17").Value = "Data Submitted Successfully. " & Now()



End Sub



 