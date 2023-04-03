Sub Send_Sheet_As_Email()

    'Declare Outlook Application object
    Dim oOutlook As Object
    Set oOutlook = CreateObject("Outlook.Application")
    
    'Create a new mail item in Outlook
    Dim oEmail As Object
    Set oEmail = oOutlook.CreateItem(olMailItem)
    
    Dim wbName As String
    wbName = ThisWorkbook.Name 'Get the workbook name
    searchText = "Accessibility testing"
    
    Dim getTotal As Integer
    Dim Critical As Integer
    Dim High As Integer
    Dim Medium As Integer
    Dim Low As Integer
    Dim chartNames(4) As String
    chartNames(0) = "WCAG Rule Wise Defect Distribution Chart"
    chartNames(1) = "Severity / Impact Wise Defect Distribution"
    chartNames(2) = "Conformance Level Wise Defect Distribution"
    chartNames(3) = "Category Wise Issue Distribution"
    chartNames(4) = "WCAG 2.1 AA Checkpoint wise Status Distribution"
    
    Critical = Worksheets("Data & Chart").Range("D60").Value
    High = Worksheets("Data & Chart").Range("E60").Value
    Medium = Worksheets("Data & Chart").Range("F60").Value
    Low = Worksheets("Data & Chart").Range("G60").Value
    getTotal = Critical + High + Medium + Low
    
    Dim htmlTable As String
    htmlTable = "<table style='border-collapse: collapse;'>"
    htmlTable = htmlTable & "<tr style='background-color: #B4C6E7; font-weight: bold;'><th style='padding: 5px; border: 1px solid black;'>Severity / Impact</th><th style='padding: 5px; border: 1px solid black;'>Definition</th></tr>"
    htmlTable = htmlTable & "<tr><td style='padding: 5px; border: 1px solid black;'>Sev 1 / Blocker</td><td style='padding: 5px; border: 1px solid black;'>The issue blocks a user from completing a core user task and there is no workaround.  This issue is a showstopper and constitutes a ship blocking bug.  Must be fixed immediately.</td></tr>"
    htmlTable = htmlTable & "<tr><td style='padding: 5px; border: 1px solid black;'>Sev 2 / High</td><td style='padding: 5px; border: 1px solid black;'>This issue isn't blocking a core user task, but it is blocking a non-core user task.  Remediation action is needed ASAP but not a ship stopper.</td></tr>"
    htmlTable = htmlTable & "<tr><td style='padding: 5px; border: 1px solid black;'>Sev 3 / Medium</td><td style='padding: 5px; border: 1px solid black;'>Remediation action is generally required in the next major release or the next time the site is updated, whichever occurs first but having less impact on the end users.</td></tr>"
    htmlTable = htmlTable & "<tr><td style='padding: 5px; border: 1px solid black;'>Sev 4 / Low</td><td style='padding: 5px; border: 1px solid black;'>This break the WCAG 2.1 checkpoints but does not affect many users or is a hindrance but not a major issue. For instance, decorative images focused and read blank.</td></tr>"
    htmlTable = htmlTable & "</table>"
    
    Dim mailBody As String
    mailBody = "<h1 style='text-align:center'>" & "End of Test Pass Report " & wbName & " Accessibility Testing</h1>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>Objectives</h2>"
    mailBody = mailBody & "<ul>"
    mailBody = mailBody & "<li> This report describes the conformance of the " & wbName & " with the W3C Web Content Accessibility Guidelines (WCAG) 2.1.</li>"
    mailBody = mailBody & "<li>Please note that this report provides an assessment of conformance to WCAG 2.1 AA Conformance Level. The results are not a certification of Compliance.</li>"
    mailBody = mailBody & "</ul>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>Key Highlights</h2>"
    mailBody = mailBody & "<ul>"
    mailBody = mailBody & "<li>A11y (Accessibility) CoE team completed the execution of " & wbName & " application on Web (dWeb+mWeb) having unique pages &amp; flows.</li>"
    mailBody = mailBody & "<li>" & wbName & " application does not meet WCAG 2.1 AA Conformance Level as many functionalities are not accessible by the End users including various techniques. </li>"
    mailBody = mailBody & "<li>Defect logging completed for all pages in the attached execution sheet with detailed descriptions along with the steps to reproduce and ready for Team review.</li>"
    mailBody = mailBody & "<li>Team logged total " & getTotal & " number of issues across all pages &amp; categorized into below categories:</li>"
    mailBody = mailBody & "<ul>"
    mailBody = mailBody & "<li>Critical Impact: " & Critical & "</li>"
    mailBody = mailBody & "<li>High Impact: " & High & "</li>"
    mailBody = mailBody & "<li>Medium Impact: " & Medium & "</li>"
    mailBody = mailBody & "<li>Low Impact: " & Low & "</li>"
    mailBody = mailBody & "</ul>"
    mailBody = mailBody & "<li>Experience Summary: " & wbName & " application is not having a good experience for the end users because users who relying over the Keyboard or Screen Readers, will face difficulties to complete the task.</li>"
    mailBody = mailBody & "<li style='font-weight:bold; text-decoration:underline; color:#843c0c;'>Key Challenges Faced by Low Vision Users:</li>"
    mailBody = mailBody & "<ul>"
    mailBody = mailBody & "<li>Dummy list point 1</li>"
    mailBody = mailBody & "<li>Dummy list point 2</li>"
    mailBody = mailBody & "</ul>"
    mailBody = mailBody & "<li style='font-weight:bold; text-decoration:underline; color:#843c0c;'>Key Challenges Faced by Keyboard Users:</li>"
    mailBody = mailBody & "<ul>"
    mailBody = mailBody & "<li>Dummy list point 1</li>"
    mailBody = mailBody & "<li>Dummy list point 2</li>"
    mailBody = mailBody & "</ul>"
    mailBody = mailBody & "<li style='font-weight:bold; text-decoration:underline; color:#843c0c;'>Key Challenges Faced by Screen Reader Users:</li>"
    mailBody = mailBody & "<ul>"
    mailBody = mailBody & "<li>Dummy list point 1</li>"
    mailBody = mailBody & "<li>Dummy list point 2</li>"
    mailBody = mailBody & "</ul>"
    mailBody = mailBody & "</ul>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>Testing Methodology</h2>"
    mailBody = mailBody & "<ul>"
    mailBody = mailBody & "<li>The " & wbName & " web application was tested against each of the applicable checkpoints on Web environment (dWeb + mWeb)</li>"
    mailBody = mailBody & "<li>The application was tested using screen readers (default setting), keyboard, Automation Accessibility Extension tools, color contrast, visual &amp; zoom.</li>"
    mailBody = mailBody & "<li>The main controls used for navigation are arrow keys, tab keys, and the H key (to skip through headings), using shift with the key combinations to reverse the flow on desktop web and for mobile web, swipe navigation (Swipe Left &amp; Swipe Right), touch exploration.</li>"
    mailBody = mailBody & "</ul>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>Execution Summary Status</h2>"
    mailBody = mailBody & "<ul>"
    mailBody = mailBody & "<li>Status: </li>"
    mailBody = mailBody & "<li>Execution Completion Rate: </li>"
    mailBody = mailBody & "</ul>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>Defect Summary Impact Wise</h2>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>Defect Summary Conformance Level Wise</h2>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>WCAG 2.1 AA Success Criteria Status Result</h2>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>WCAG Failure by Rules: </h2>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>WCAG Rule Wise Defect Distribution Chart</h2>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>Severity / Impact Wise Defect Distribution</h2>"
    mailBody = mailBody & "<br><h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>Conformance Level Wise Defect Distribution</h2>"
    mailBody = mailBody & "<br><h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>Category Wise Issue Distribution</h2>"
    mailBody = mailBody & "<br><h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>WCAG 2.1 AA Checkpoint wise Status Distribution</h2>"
    mailBody = mailBody & "<br><h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>Test Environment Summary</h2>"
    mailBody = mailBody & "<ul>"
    mailBody = mailBody & "<li>N/A</li>"
    mailBody = mailBody & "</ul>"
    mailBody = mailBody & "<h2 style='margin-top:2pt; margin-right:0in; margin-bottom:0in; font-size:12pt; font-family:Calibri Light, Sans-serif; color:#2f5496; font-weight:normal;'>References</h2>"
    mailBody = mailBody & "<ul>"
    mailBody = mailBody & "<li>Web Content Accessibility Guideline Documentation: WCAG 2.1</li>"
    mailBody = mailBody & "<li>MFB Bank OLE Web Application – </li>"
    mailBody = mailBody & "<li>Severity/Impact of the Defects is defined based on:</li>"
    mailBody = mailBody & htmlTable
    mailBody = mailBody & "</ul>"
     

    If InStr(1, LCase(wbName), LCase(searchText)) > 0 Then
        wbName = Replace(wbName, searchText, "")
    End If
    
    

    With oEmail
        'Set recipient, subject, and body of the email
        .To = "a@gmail.com"
        .Subject = "This is a test email"
        .BodyFormat = olFormatHTML
        .HTMLBody = mailBody
        'Display the email
        .Display
        
        'Declare Inspector and Word Document objects
        Dim oOutlookInscept As Outlook.Inspector
        Dim oWordDoc As Word.Document
        Dim oChartobj As ChartObject


        'Get the inspector for the email and retrieve the Word Document object
        Set oOutlookInspect = .GetInspector
        Set oWordDocl = oOutlookInspect.WordEditor
        
        
'.............................................................................................Status_Logging_Table_Range Code Started .............................................................................................
        'Store the Status_Logging_Table range including the header row in a variable
        Dim Status_Logging_Table_Range As Range
        Set Status_Logging_Table_Range = ActiveSheet.Range("Status_Logging_Table[#Headers]").Resize(ActiveSheet.Range("Status_Logging_Table").Rows.Count + 1)
        
        'Copy the Status_Logging_Table table including the header row to the email body
        Status_Logging_Table_Range.Copy
        
        'Insert the Status_Logging_Table table including the header row into the email body
        Set oWordRng2 = oWordDocl.Application.ActiveDocument.Content
        oWordRng2.Find.Execute FindText:="Execution Completion Rate: "
        oWordRng2.Collapse Direction:=wdCollapseEnd
        oWordRng2.InsertAfter " " & vbNewLine
        oWordRng2.Paste
        
'.............................................................................................Status_Logging_Table_Range Code Ended .............................................................................................
        
        
'.............................................................................................Defect_Logging_Table_Range Code Started ...................................................................................
        'Store the Defect_Logging_Table range including the header row in a variable
        Dim Defect_Logging_Table_Range As Range
        Set Defect_Logging_Table_Range = ActiveSheet.Range("Defect_Logging_Table[#Headers]").Resize(ActiveSheet.Range("Defect_Logging_Table").Rows.Count + 1)
        'Copy the Defect_Logging_Table table including the header row to the email body
        Defect_Logging_Table_Range.Copy
        
        'Insert the Defect_Logging_Table table including the header row into the email body
        Set oWordRng3 = oWordDocl.Application.ActiveDocument.Content
        oWordRng3.Find.Execute FindText:="Defect Summary Impact Wise"
        oWordRng3.Collapse Direction:=wdCollapseEnd
        oWordRng3.Paste
        
'.............................................................................................Defect_Logging_Table_Range Code Ended ...................................................................................
        
               
        
'.............................................................................................Conf_Logging_Table Code Started ...................................................................................
        'Store the Conf_Logging_Table range including the header row in a variable
        Dim Conf_Logging_Table_Range As Range
        Set Conf_Logging_Table_Range = ActiveSheet.Range("Conf_Logging_Table[#Headers]").Resize(ActiveSheet.Range("Conf_Logging_Table").Rows.Count + 1)
        
        'Copy the Conf_Logging_Table table including the header row to the email body
        Conf_Logging_Table_Range.Copy
        
        'Insert the Conf_Logging_Table table including the header row into the email body
        Set oWordRng4 = oWordDocl.Application.ActiveDocument.Content
        oWordRng4.Find.Execute FindText:="Defect Summary Conformance Level Wise"
        oWordRng4.Collapse Direction:=wdCollapseEnd
        oWordRng4.Paste
        
'.............................................................................................Conf_Logging_Table Code Ended ...................................................................................

'.....................................................................................................WCAG Table Code Started .............................................................................................
        
        'Store the WCAG_Counts range including the header row in a variable
        Dim WCAG_Counts_Range As Range
        Set WCAG_Counts_Range = ActiveSheet.Range("WCAG_Counts[#Headers]").Resize(ActiveSheet.Range("WCAG_Counts").Rows.Count + 1)
        
        'Copy the WCAG_Counts table including the header row to the email body
        WCAG_Counts_Range.Copy
        
        'Insert the WCAG_Counts table including the header row into the email body
        Set oWordRng5 = oWordDocl.Application.ActiveDocument.Content
        oWordRng5.Find.Execute FindText:="WCAG 2.1 AA Success Criteria Status Result"
        oWordRng5.Collapse Direction:=wdCollapseEnd
        oWordRng5.Paste
        
'...................................................................................................WCAG Table Code Ended .............................................................................................
               
               
               
'.....................................................................................................WCAG Table Code Started .............................................................................................
        
        'Store the WCAG_Counts range including the header row in a variable
        Dim Status_table_Range As Range
        Set Status_table_Range = ActiveSheet.Range("Status_table[#Headers]").Resize(ActiveSheet.Range("Status_table").Rows.Count + 1)
        
        'Copy the WCAG_Counts table including the header row to the email body
        Status_table_Range.Copy
        
        'Insert the WCAG_Counts table including the header row into the email body
        Set oWordRng6 = oWordDocl.Application.ActiveDocument.Content
        oWordRng6.Find.Execute FindText:="WCAG Failure by Rules: "
        oWordRng6.Collapse Direction:=wdCollapseEnd
        oWordRng6.Paste
        
'...................................................................................................WCAG Table Code Ended .............................................................................................
        Dim trav As Integer
        trav = 0
        
        'Loop through all chart objects in the active sheet and paste them into the email body
        For Each oChartobj In ActiveSheet.ChartObjects
            oChartobj.Chart.ChartArea.Copy
        
            'Get the inspector for the email and retrieve the Word Document object
            Set oOutlookInscept = .GetInspector
            Set oWordDoc = oOutlookInscept.WordEditor
        
            'Insert the chart into the email body
            Set oWordRng = oWordDoc.Application.ActiveDocument.Content
            oWordRng.Find.Execute FindText:=chartNames(trav)
            oWordRng.Collapse Direction:=wdCollapseEnd
            oWordRng.InsertParagraphAfter
            oWordRng.Collapse Direction:=wdCollapseEnd
            oWordRng.Paste
            trav = trav + 1
        Next
        


    
    End With
    
    'Release the email and Outlook objects from memory
    Set oEmail = Nothing
    Set oOutlook = Nothing
    
End Sub
