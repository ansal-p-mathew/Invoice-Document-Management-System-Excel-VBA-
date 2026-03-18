Attribute VB_Name = "Module1"
Sub UploadDocuments()

Dim ws As Worksheet
Dim rowNum As Long
Dim BasePath As String
Dim InvoiceNumber As String
Dim InvoiceFolder As String
Dim sep As String
Dim fd As FileDialog
Dim selectedFile As String
Dim FileName As String
Dim i As Long
Dim totalFiles As Long

Set ws = ActiveSheet

If ActiveCell.Row = 1 Then
MsgBox "Please select an invoice row first.", vbExclamation
Exit Sub
End If

rowNum = ActiveCell.Row
sep = Application.PathSeparator

InvoiceNumber = Trim(ws.Cells(rowNum, 4).Value)

If InvoiceNumber = "" Then
MsgBox "Enter Customer Invoice first.", vbExclamation
Exit Sub
End If

BasePath = Trim(ws.Range("K2").Value)

If Right(BasePath, 1) <> sep Then BasePath = BasePath & sep

InvoiceFolder = BasePath & InvoiceNumber

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd

.AllowMultiSelect = True
.Title = "Upload Documents"

If .Show = -1 Then

For i = 1 To .SelectedItems.Count

selectedFile = .SelectedItems(i)
FileName = Mid(selectedFile, InStrRev(selectedFile, sep) + 1)

FileCopy selectedFile, InvoiceFolder & sep & FileName

Next i

End If

End With

totalFiles = CountFiles(InvoiceFolder)

ws.Cells(rowNum, 6).Value = totalFiles & " Doc" & IIf(totalFiles > 1, "s", "")
ws.Cells(rowNum, 6).HorizontalAlignment = xlCenter
ws.Cells(rowNum, 6).Font.Bold = (totalFiles > 0)

End Sub

Sub UploadReceipts()

Dim ws As Worksheet
Dim rowNum As Long
Dim BasePath As String
Dim InvoiceNumber As String
Dim ReceiptFolder As String
Dim sep As String
Dim fd As FileDialog
Dim selectedFile As String
Dim FileName As String
Dim i As Long
Dim totalFiles As Long

Set ws = ActiveSheet

If ActiveCell.Row = 1 Then
MsgBox "Please select an invoice row first.", vbExclamation
Exit Sub
End If

rowNum = ActiveCell.Row
sep = Application.PathSeparator

InvoiceNumber = Trim(ws.Cells(rowNum, 4).Value)

If InvoiceNumber = "" Then
MsgBox "Enter Customer Invoice first.", vbExclamation
Exit Sub
End If

BasePath = Trim(ws.Range("K2").Value)

If Right(BasePath, 1) <> sep Then BasePath = BasePath & sep

ReceiptFolder = BasePath & InvoiceNumber & sep & "Payment Receipts" & sep

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd

.AllowMultiSelect = True
.Title = "Upload Receipts"

If .Show = -1 Then

For i = 1 To .SelectedItems.Count

selectedFile = .SelectedItems(i)
FileName = Mid(selectedFile, InStrRev(selectedFile, sep) + 1)

FileCopy selectedFile, ReceiptFolder & FileName

Next i

End If

End With

totalFiles = CountFiles(ReceiptFolder)

ws.Cells(rowNum, 7).Value = totalFiles & " Receipt" & IIf(totalFiles > 1, "s", "")
ws.Cells(rowNum, 7).HorizontalAlignment = xlCenter
ws.Cells(rowNum, 7).Font.Bold = (totalFiles > 0)

End Sub


Function CountFiles(folderPath As String) As Long

Dim f As String
Dim sep As String

sep = Application.PathSeparator

f = Dir(folderPath & sep & "*")

Do While f <> ""
    CountFiles = CountFiles + 1
    f = Dir
Loop

End Function
