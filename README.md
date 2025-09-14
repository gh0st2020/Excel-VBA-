<h1>EXCEL VBA - Introduced to work place </h1>

<h2>Description</h2>

<img width="1134" height="735" alt="image" src="https://github.com/user-attachments/assets/1a2ffd8f-9ec9-4504-869f-ce8e00aa21d6" />


The Hedge Fund Order Tracker created by me is a macro‑enabled Excel solution designed to standardise and control the end‑to‑end workflow for hedge fund subscriptions and redemptions. It captures each order with a unique reference, records the direction as Subscription or Redemption, stamps the date the order was sent, and guides follow‑up actions through a colour‑coded status and an automated prompt. 

At the core of the tracker are seven columns that mirror the actual operational steps. “Ref” serves as the order identifier and preserves business formats (for example, SUB‑2025‑001). “Direction” allows the user to type “S” or “R,” which automatically expands to Subscription or Redemption and applies an intuitive cell colour so the blotter is readable at a glance. “Date Order Sent” is populated automatically when a reference is entered, ensuring a reliable audit trail in dd/mm/yyyy format. “Status” is selected from a controlled drop‑down list—Order Sent, TA acknowledge email, Processing / AML, and Trade Date Confirmed—with each choice applying a distinct fill colour (red, orange, yellow, green) to make progress immediately visible. “Last Contacted” records the most recent date of outreach to the transfer agent in dd/mm/yyyy format. “Action” displays a prominent “Confirm Trade Date” message in large, bold red text whenever the status is still in an early stage (Order Sent, TA acknowledge email, or Processing / AML) and at least two days have elapsed since the last contact. Finally, “Order Cut Off” holds the fund’s deadline date, so the team can monitor proximity to critical milestones.

This project was undertaken to address persistent operational pain points in the subscription/redemption process. Teams often rely on email threads to infer whether the transfer agent has received the documents, commenced AML checks, or confirmed a trade date. Because hedge funds typically deal on monthly or quarterly cycles with strict cut‑off dates, a missed follow‑up can lead to delayed execution, manual rework, and client dissatisfaction. The tracker reduces these risks by centralising the information that matters, enforcing consistent status definitions, and automatically prompting timely chasers. By shifting from ad‑hoc email monitoring to a structured, time‑aware blotter, the solution lowers the cognitive load on busy dealers, improves transparency for the desk and middle office, and strengthens the audit trail for reviews.

Introducing the tracker into the workplace is deliberately low‑friction. The file is an Excel workbook that runs on standard desktop builds, so users do not need new logins or platforms. The recommended rollout is to publish a read‑only master on SharePoint or Teams with versioning enabled and to designate a small pilot group for one to two weeks to validate the status list, colour conventions, and chase timing. A brief training (about 30 minutes) is typically sufficient: users learn to enter a new reference, select the appropriate status, update the last‑contacted date after each TA interaction, and respond to the Action prompt when it appears. Post‑pilot, the desk can formalise a short operating procedure describing who maintains the tracker, how exceptions are escalated, and how often KPIs are reviewed.

The operational impact is immediate. Date stamping on entry provides consistent evidence of when orders were sent. The controlled status list eliminates ambiguous phrasing and improves hand‑offs between dealers, middle office, and management. The Action prompt ensures that follow‑ups occur on a predictable rhythm, reducing the likelihood of missing cut‑offs. Because the logic uses standard Excel functions (such as TODAY()) and light event‑based VBA, the “Confirm Trade Date” message updates automatically as time passes, without users having to refresh or run scripts. The solution also prepares the desk for basic metrics: time from Order Sent to TA acknowledgement, time to trade date confirmation, and counts of items approaching cut‑off, all of which can be exported or connected to Power BI if desired.

In summary, the Hedge Fund Order Tracker formalises a previously manual, email‑driven process into a clear, colour‑coded workflow with built‑in reminders. It was created by me to reduce the operational risk of missed cut‑offs, improve team coordination, and provide an auditable view of order progress, while requiring minimal change for users. Its introduction into the workplace is straightforward, cost‑effective, and immediately beneficial to service quality and control.
<br />

<img width="1708" height="1024" alt="image" src="https://github.com/user-attachments/assets/1d9c7427-4e0d-4c59-91e5-ace0249077be" />

<img width="1903" height="1039" alt="image" src="https://github.com/user-attachments/assets/816a4369-5f03-4045-a85a-297d1e1fc0e4" />


<h2>Code</h2>

MODULE: 
Option Explicit

Sub Setup_HF_Tracker()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim fc As FormatCondition
    Dim fml As String

    Set ws = ThisWorkbook.Worksheets("Sheet1") ' <-- change if needed
    LastRow = 1000                              ' prep rows 2:1000

    '--- Headers
    ws.Range("A1").Value = "Ref"
    ws.Range("B1").Value = "Direction"
    ws.Range("C1").Value = "Date Order Sent"
    ws.Range("D1").Value = "Status"
    ws.Range("E1").Value = "Last Contacted"    ' << renamed per your request
    ws.Range("F1").Value = "Action"
    ws.Range("G1").Value = "Order Cut Off"

    '--- Formats
    ws.Columns("A").NumberFormat = "@"                 ' keep IDs like 00123
    ws.Columns("C").NumberFormat = "dd/mm/yyyy"
    ws.Columns("E").NumberFormat = "dd/mm/yyyy"
    ws.Columns("G").NumberFormat = "dd/mm/yyyy"

    ws.Range("A1:G1").Font.Bold = True
    ws.Columns("A:G").AutoFit

    ' Freeze top row
    ws.Activate
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True

    '--- STATUS (D): Data Validation list (updated labels)
    With ws.Range("D2:D" & LastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="Order Sent,TA acknowledge email,Processing / AML,Trade Date Confirmed"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    '--- STATUS (D): Conditional fill colours
    ws.Range("D2:D" & LastRow).FormatConditions.Delete

    Set fc = ws.Range("D2:D" & LastRow).FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=$D2=""Order Sent""")
    fc.Interior.Color = RGB(255, 199, 206)    ' light red

    Set fc = ws.Range("D2:D" & LastRow).FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=$D2=""TA acknowledge email""")
    fc.Interior.Color = RGB(255, 217, 102)    ' orange

    Set fc = ws.Range("D2:D" & LastRow).FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=$D2=""Processing / AML""")
    fc.Interior.Color = RGB(255, 242, 204)    ' yellow

    Set fc = ws.Range("D2:D" & LastRow).FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=$D2=""Trade Date Confirmed""")
    fc.Interior.Color = RGB(198, 239, 206)    ' green

    '--- E & G: Date validation (basic sanity check)
    With ws.Range("E2:E" & LastRow).Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=DATE(2000,1,1)", Formula2:="=DATE(2099,12,31)"
        .IgnoreBlank = True
        .InputTitle = "Date (dd/mm/yyyy)"
        .InputMessage = "Type a valid date (dd/mm/yyyy)."
    End With
    With ws.Range("G2:G" & LastRow).Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=DATE(2000,1,1)", Formula2:="=DATE(2099,12,31)"
        .IgnoreBlank = True
        .InputTitle = "Date (dd/mm/yyyy)"
        .InputMessage = "Type a valid date (dd/mm/yyyy)."
    End With

    '--- F: Action formula (auto-updates daily via TODAY())
    ' Show "Confirm Trade Date" when:
    '   D is Order Sent OR TA acknowledge email OR Processing / AML
    '   AND E is a valid date
    '   AND (TODAY() - E) >= 2
    fml = "=IF(AND(COUNT($E2)=1,OR($D2=""Order Sent"",$D2=""TA acknowledge email"",$D2=""Processing / AML""),TODAY()-$E2>=2),""Confirm Trade Date"","""")"
    ws.Range("F2:F" & LastRow).Formula = fml

    ' (No CF on F; a Calculate event will set big bold red only when needed)

    MsgBox "Tracker updated: D list/colours, E renamed, F logic set.", vbInformation
End Sub

Sheet1:
Option Explicit

' Fires when cells change (typing/paste)
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    Application.EnableEvents = False

    Dim rng As Range, cell As Range
    Dim val As String

    ' 1) Ref entered in Column A -> date stamp Column C (dd/mm/yyyy)
    Set rng = Intersect(Target, Me.Columns("A"))
    If Not rng Is Nothing Then
        For Each cell In rng.Cells
            If cell.Row > 1 Then
                If Len(cell.Value) > 0 Then
                    If Len(Me.Cells(cell.Row, "C").Value) = 0 Then
                        Me.Cells(cell.Row, "C").Value = Date
                        Me.Cells(cell.Row, "C").NumberFormat = "dd/mm/yyyy"
                    End If
                    ' Ensure F's formula exists on that row
                    If Len(Me.Cells(cell.Row, "F").Formula) = 0 Then
                        Me.Cells(cell.Row, "F").FormulaR1C1 = _
                            "=IF(AND(COUNT(RC[-1])=1,OR(RC[-2]=""Order Sent"",RC[-2]=""TA acknowledge email"",RC[-2]=""Processing / AML""),TODAY()-RC[-1]>=2),""Confirm Trade Date"","""")"
                    End If
                Else
                    Me.Cells(cell.Row, "C").ClearContents
                End If
            End If
        Next cell
    End If

    ' 2) Direction in Column B -> normalize S/R + colour fill
    Set rng = Intersect(Target, Me.Columns("B"))
    If Not rng Is Nothing Then
        For Each cell In rng.Cells
            If cell.Row > 1 Then
                val = Trim$(UCase$(CStr(cell.Value)))
                Select Case val
                    Case "S", "SUBSCRIPTION"
                        cell.Value = "Subscription"
                        cell.Interior.Color = RGB(198, 239, 206)  ' light green
                    Case "R", "REDEMPTION"
                        cell.Value = "Redemption"
                        cell.Interior.Color = RGB(255, 199, 206)  ' light red
                    Case ""
                        cell.Interior.Pattern = xlNone
                    Case Else
                        MsgBox "Enter 'S' for Subscription or 'R' for Redemption.", vbExclamation, "Invalid Direction"
                        cell.ClearContents
                        cell.Interior.Pattern = xlNone
                End Select
            End If
        Next cell
    End If

CleanExit:
    Application.EnableEvents = True
End Sub

' Fires on any calculation (TODAY() ticks daily, or manual recalc)
Private Sub Worksheet_Calculate()
    Dim rng As Range, c As Range
    Set rng = Me.Range("F2:F1000") ' keep in sync with LastRow in setup

    For Each c In rng.Cells
        If c.Value = "Confirm Trade Date" Then
            With c.Font
                .Bold = True
                .Color = RGB(192, 0, 0)   ' dark red for visibility
                .Size = 14                ' "large" font
            End With
        Else
            With c.Font
                .Bold = False
                .ColorIndex = xlAutomatic
                .Size = 11                ' normal font
            End With
        End If
    Next c
End Sub







