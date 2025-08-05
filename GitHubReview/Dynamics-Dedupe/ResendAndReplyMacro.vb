Sub ResendAndReplyMacro()
    ' Resend and Reply Macro for RCI Talent Sourcing Outreach
    ' This macro processes selected emails to either resend or reply with the talent sourcing message
    
    Dim olApp As Outlook.Application
    Dim olSelection As Outlook.Selection
    Dim olMail As Outlook.MailItem
    Dim olNewMail As Outlook.MailItem
    Dim i As Integer
    Dim sentCount As Integer
    Dim duplicateCount As Integer
    Dim errorCount As Integer
    Dim recipientEmail As String
    Dim recipientName As String
    Dim messageBody As String
    Dim isReply As Boolean
    Dim originalDisplayAlerts As Boolean
    Dim startTime As Date
    Dim processedEmails As Object ' Dictionary to track processed email addresses
    Dim userSignature As String ' Cache signature to avoid repeated file reads
    
    ' Initialize Outlook application
    Set olApp = Application
    Set olSelection = olApp.ActiveExplorer.Selection
    
    ' Check if any items are selected
    If olSelection.Count = 0 Then
        MsgBox "Please select one or more emails to process.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' Initialize variables
    Set processedEmails = CreateObject("Scripting.Dictionary")
    processedEmails.CompareMode = 1 ' Case-insensitive comparison
    sentCount = 0
    duplicateCount = 0
    errorCount = 0
    
    ' Get user signature once (performance optimization)
    userSignature = GetOutlookSignature()
    
    ' Record start time and show initial message
    startTime = Now
    MsgBox "Starting bulk email processing..." & vbCrLf & "Selected emails: " & olSelection.Count & vbCrLf & vbCrLf & "The macro will automatically skip duplicate recipients." & vbCrLf & "You can continue working in Outlook while it processes.", vbInformation, "Macro Started"
    
    ' Store original display alerts setting and turn off alerts
    originalDisplayAlerts = olApp.DisplayAlerts
    olApp.DisplayAlerts = False
    
    ' Your message template
    messageBody = "Need more qualified, interested and local candidate" & vbCrLf & vbCrLf & _
                 ". Leverage the latest in Ai candidate sourcing technology" & vbCrLf & vbCrLf & _
                 "RCI has just recently implemented the very latest in Ai and Machine Learning technology that has greatly enhanced our proprietary Talent Locator solution. Much more than just posting your jobs or sourcing indeed and LinkedIn. Now your team can have access to over 300 million candidates without the swipe of a mouse. We support any roles (exempt or non exempt)" & vbCrLf & vbCrLf & _
                 "Let's speak this week/month about your current needs and challenges…whether 1 or 1000 openings Talentlocator has a scalable model that you can turn on or off." & vbCrLf & vbCrLf & _
                 "Our clients are seeing dramatic results."
    
    ' Process each selected email
    For i = 1 To olSelection.Count
        On Error Resume Next ' Handle errors gracefully
        
        If TypeName(olSelection.Item(i)) = "MailItem" Then
            Set olMail = olSelection.Item(i)
            
            ' Determine if this should be a reply or resend
            ' If the current user is the sender, it's a resend; otherwise, it's a reply
            isReply = Not IsCurrentUserSender(olMail)
            
            If isReply Then
                ' REPLY: Get sender's email and name from the original email
                recipientEmail = olMail.SenderEmailAddress
                recipientName = ExtractFirstName(olMail.SenderName)
            Else
                ' RESEND: Get recipient's email and name from the original email's To field
                Dim recipientInfo As String
                recipientInfo = GetFirstExternalRecipientWithName(olMail)
                If recipientInfo <> "" Then
                    ' Parse email and name from the returned string
                    Dim parts() As String
                    parts = Split(recipientInfo, "|")
                    recipientEmail = parts(0)
                    If UBound(parts) > 0 Then
                        recipientName = ExtractFirstName(parts(1))
                    Else
                        recipientName = ExtractFirstName("")
                    End If
                End If
            End If
            
            ' Only proceed if we have a valid external recipient
            If recipientEmail <> "" And Not IsInternalEmail(recipientEmail) Then
                
                ' Check for duplicate email address
                If processedEmails.Exists(LCase(recipientEmail)) Then
                    ' This is a duplicate - skip it
                    duplicateCount = duplicateCount + 1
                Else
                    ' Add email to processed list
                    processedEmails.Add LCase(recipientEmail), True
                    
                    ' Create new email
                    Set olNewMail = olApp.CreateItem(olMailItem)
                    
                    With olNewMail
                        .To = recipientEmail
                        .Subject = IIf(isReply, "RE: " & olMail.Subject, "Enhanced AI Candidate Sourcing Solutions")
                        
                        ' Build email body with conditional greeting
                        If recipientName <> "there" And recipientName <> "" Then
                            .Body = "Dear " & recipientName & "," & vbCrLf & vbCrLf & messageBody
                        Else
                            .Body = messageBody
                        End If
                        
                        ' Get signature and append it
                        Dim signature As String
                        signature = GetOutlookSignature()
                        If signature <> "" Then
                            .HTMLBody = .HTMLBody & "<br><br>" & signature
                        End If
                        
                        ' Send silently without displaying
                        .Send
                    End With
                    
                    sentCount = sentCount + 1
                End If
                
                ' Close the original email if it's open
                If olMail.GetInspector.IsWordMail Then
                    olMail.Close olSave
                End If
            End If
        End If
        
        On Error GoTo 0 ' Reset error handling
    Next i
    
    ' Restore original display alerts setting
    olApp.DisplayAlerts = originalDisplayAlerts
    
    ' Show completion message
    MsgBox "Macro completed successfully!" & vbCrLf & "Emails sent: " & sentCount, vbInformation, "Process Complete"
    
    ' Clean up
    Set processedEmails = Nothing
    Set olNewMail = Nothing
    Set olMail = Nothing
    Set olSelection = Nothing
    Set olApp = Nothing
End Sub

Function IsCurrentUserSender(olMail As Outlook.MailItem) As Boolean
    ' Check if the current user is the sender of the email
    ' Optimized for Exchange Online/Office 365
    On Error Resume Next
    Dim currentUserEmail As String
    Dim currentUser As Outlook.AddressEntry
    
    ' Get current user's primary SMTP address from Exchange
    Set currentUser = Application.Session.CurrentUser.AddressEntry
    currentUserEmail = currentUser.GetExchangeUser.PrimarySmtpAddress
    
    ' Compare sender email with current user email
    IsCurrentUserSender = (LCase(olMail.SenderEmailAddress) = LCase(currentUserEmail))
    
    On Error GoTo 0
End Function

Function CreateAndSendEmail(olApp As Outlook.Application, recipientEmail As String, recipientName As String, messageBody As String, isReply As Boolean, originalSubject As String, userSignature As String) As Boolean
    ' Create and send email with proper error handling
    ' Returns True if successful, False if failed
    
    On Error Resume Next
    
    Dim olNewMail As Outlook.MailItem
    Dim emailBody As String
    Dim emailSubject As String
    
    ' Create new email
    Set olNewMail = olApp.CreateItem(olMailItem)
    
    ' Build email subject
    emailSubject = IIf(isReply, "RE: " & originalSubject, "Enhanced AI Candidate Sourcing Solutions")
    
    ' Build email body with conditional greeting
    If recipientName <> "there" And recipientName <> "" Then
        emailBody = "Dear " & recipientName & "," & vbCrLf & vbCrLf & messageBody
    Else
        emailBody = messageBody
    End If
    
    ' Convert to HTML and add signature
    emailBody = Replace(emailBody, vbCrLf, "<br>")
    If userSignature <> "" Then
        emailBody = emailBody & "<br><br>" & userSignature
    End If
    
    With olNewMail
        .To = recipientEmail
        .Subject = emailSubject
        .HTMLBody = emailBody
        .Send
    End With
    
    ' Check if send was successful
    CreateAndSendEmail = (Err.Number = 0)
    
    ' Clean up
    Set olNewMail = Nothing
    On Error GoTo 0
End Function

Function ExtractFirstName(fullName As String) As String
    ' Extract first name from full name
    Dim names() As String
    Dim firstName As String
    
    ' Remove common email patterns and clean the name
    fullName = Replace(fullName, "'", "")
    fullName = Trim(fullName)
    
    ' Split by space and take the first part
    names = Split(fullName, " ")
    firstName = names(0)
    
    ' Clean up common prefixes
    If LCase(firstName) = "mr." Or LCase(firstName) = "ms." Or LCase(firstName) = "mrs." Or LCase(firstName) = "dr." Then
        If UBound(names) > 0 Then
            firstName = names(1)
        End If
    End If
    
    ' If still empty or looks like an email, use a generic greeting
    If firstName = "" Or InStr(firstName, "@") > 0 Then
        firstName = "there"
    End If
    
    ExtractFirstName = firstName
End Function

Function ExtractFirstNameFromEmail(olMail As Outlook.MailItem, emailAddress As String) As String
    ' Extract first name from the recipient list based on email address
    ' Optimized for Exchange Online/Office 365
    Dim recipient As Outlook.recipient
    Dim i As Integer
    
    On Error Resume Next
    
    For i = 1 To olMail.Recipients.Count
        Set recipient = olMail.Recipients.Item(i)
        If LCase(recipient.AddressEntry.GetExchangeUser.PrimarySmtpAddress) = LCase(emailAddress) Then
            ExtractFirstNameFromEmail = ExtractFirstName(recipient.Name)
            Exit Function
        End If
    Next i
    
    ' If not found, extract from email address
    Dim emailParts() As String
    emailParts = Split(emailAddress, "@")
    If UBound(emailParts) >= 0 Then
        ExtractFirstNameFromEmail = ExtractFirstName(emailParts(0))
    Else
        ExtractFirstNameFromEmail = "there"
    End If
    
    On Error GoTo 0
End Function

Function GetFirstExternalRecipient(olMail As Outlook.MailItem) As String
    ' Get the first external recipient from the To field
    ' Optimized for Exchange Online/Office 365
    Dim recipient As Outlook.recipient
    Dim i As Integer
    Dim emailAddress As String
    
    On Error Resume Next
    
    For i = 1 To olMail.Recipients.Count
        Set recipient = olMail.Recipients.Item(i)
        
        ' Get SMTP email address from Exchange
        emailAddress = recipient.AddressEntry.GetExchangeUser.PrimarySmtpAddress
        
        ' Check if it's external and valid
        If emailAddress <> "" And Not IsInternalEmail(emailAddress) Then
            GetFirstExternalRecipient = emailAddress
            Exit Function
        End If
    Next i
    
    GetFirstExternalRecipient = ""
    On Error GoTo 0
End Function

Function GetFirstExternalRecipientWithName(olMail As Outlook.MailItem) As String
    ' Return the first external recipient with both email and name separated by a pipe
    Dim recipient As Outlook.recipient
    Dim emailAddress As String
    
    On Error Resume Next
    For Each recipient In olMail.Recipients
        emailAddress = recipient.AddressEntry.GetExchangeUser.PrimarySmtpAddress
        If emailAddress <> "" And Not IsInternalEmail(emailAddress) Then
            GetFirstExternalRecipientWithName = emailAddress & "|" & recipient.Name
            Exit Function
        End If
    Next recipient
    GetFirstExternalRecipientWithName = ""
    On Error GoTo 0
End Function

Function IsInternalEmail(emailAddress As String) As Boolean
    ' Check if email address is internal (rciars.com or rcirs.com)
    emailAddress = LCase(emailAddress)
    IsInternalEmail = (InStr(emailAddress, "@rciars.com") > 0) Or (InStr(emailAddress, "@rcirs.com") > 0)
End Function

Function GetOutlookSignature() As String
    ' Get the default Outlook signature with improved reliability
    On Error Resume Next
    
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim strSignaturePath As String
    Dim strSignature As String
    Dim defaultSigName As String
    
    ' Create FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Try to get default signature name from registry first
    On Error Resume Next
    defaultSigName = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Common\MailSettings\NewSignature")
    On Error GoTo 0
    
    ' Get signature folder path
    strSignaturePath = Environ("APPDATA") & "\Microsoft\Signatures\"
    
    ' Check if signatures folder exists
    If objFSO.FolderExists(strSignaturePath) Then
        Set objFolder = objFSO.GetFolder(strSignaturePath)
        
        ' Try to find the default signature first
        If defaultSigName <> "" Then
            Dim defaultSigPath As String
            defaultSigPath = strSignaturePath & defaultSigName & ".htm"
            If objFSO.FileExists(defaultSigPath) Then
                Dim objTextStream As Object
                Set objTextStream = objFSO.OpenTextFile(defaultSigPath, 1)
                strSignature = objTextStream.ReadAll
                objTextStream.Close
                GetOutlookSignature = strSignature
                Exit Function
            End If
        End If
        
        ' Fallback: Look for any .htm file (backwards compatibility)
        For Each objFile In objFolder.Files
            If LCase(Right(objFile.Name, 4)) = ".htm" Then
                Dim objStream As Object
                Set objStream = objFSO.OpenTextFile(objFile.Path, 1)
                strSignature = objStream.ReadAll
                objStream.Close
                Exit For
            End If
        Next objFile
    End If
    
    GetOutlookSignature = strSignature
    
    On Error GoTo 0
End Function
