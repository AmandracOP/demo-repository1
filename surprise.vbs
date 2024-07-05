Do
    choice = MsgBox("Will you go on a date with me when we get back to campus?", vbYesNo + vbQuestion + vbSystemModal, "Ask for a Date")
    If choice = vbYes Then
        MsgBox "Thank you!", vbInformation + vbSystemModal, "Confirmation"
        Exit Do
    End If
Loop
