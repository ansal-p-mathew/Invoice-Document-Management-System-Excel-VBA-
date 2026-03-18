VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo SafeExit

' ?? Prevent manual edits in F & G (typing, paste)
If Not Intersect(Target, Me.Columns("F:G")) Is Nothing Then
    Application.EnableEvents = False
    Application.Undo
    MsgBox "You cannot edit this cell.", vbExclamation
    GoTo SafeExit
End If

If Target.CountLarge > 1 Then GoTo SafeExit

If Intersect(Target, Me.Columns("A")) Is Nothing And _
   Intersect(Target, Me.Columns("D")) Is Nothing Then GoTo SafeExit

Application.EnableEvents = False

Dim BasePath As String
Dim InvoiceFolder As String
Dim InvoiceNumber As String
Dim CustomerPO As String
Dim sep As String

sep = Application.PathSeparator

CustomerPO = Trim(Me.Cells(Target.Row, 1).Value)
InvoiceNumber = Trim(Me.Cells(Target.Row, 4).Value)

If CustomerPO <> "" And InvoiceNumber <> "" Then

    BasePath = Trim(Me.Range("K2").Value)

    If BasePath = "" Then GoTo SafeExit

    If Right(BasePath, 1) <> sep Then BasePath = BasePath & sep

    If Dir(BasePath, vbDirectory) = "" Then MkDir BasePath

    InvoiceFolder = BasePath & InvoiceNumber

    If Dir(InvoiceFolder, vbDirectory) = "" Then MkDir InvoiceFolder

    If Dir(InvoiceFolder & sep & "Payment Receipts", vbDirectory) = "" Then
        MkDir InvoiceFolder & sep & "Payment Receipts"
    End If

    Me.Cells(Target.Row, 6).Value = "0 Docs"
    Me.Cells(Target.Row, 7).Value = "0 Receipts"

    Me.Cells(Target.Row, 6).HorizontalAlignment = xlCenter
    Me.Cells(Target.Row, 7).HorizontalAlignment = xlCenter

    Me.Cells(Target.Row, 6).Font.Bold = False
    Me.Cells(Target.Row, 7).Font.Bold = False

End If

SafeExit:
Application.EnableEvents = True

End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

' ?? Block double-click editing in F & G
If Not Intersect(Target, Me.Columns("F:G")) Is Nothing Then
    Cancel = True
    MsgBox "You cannot edit this cell.", vbExclamation
End If

End Sub

