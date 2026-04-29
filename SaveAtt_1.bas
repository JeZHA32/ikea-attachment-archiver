Attribute VB_Name = "Module1"
'==============================================================================
' SaveAtt — Outlook VBA macro
'
' Purpose: Automate the carton-photo email triage that the IKEA Logistics
' quality team was doing by hand.
'
' What it does:
'   1. Iterates over the user-selected emails in the active Outlook explorer
'   2. For each email, parses a shipment identifier (HM) out of the subject line
'      — the identifier sits between two fixed markers ("Info DEVRPT" and "IKEA")
'   3. Creates a per-shipment folder under a configured archive root if one
'      doesn't already exist
'   4. Saves every attachment of that email into the matching shipment folder
'
' Design choices (worth flagging in interview):
'   - Operates on Selection, not the entire Inbox — the user keeps manual
'     control over which emails to archive
'   - Skips emails with no attachments (no empty folders, no error)
'   - Folder is created on demand with Dir(...) check, so re-running on the
'     same shipment is a no-op rather than an error
'
' Originally deployed at IKEA Logistics Services (Shanghai), 2017.
' This is a sanitized version — original archive root path replaced with a
' generic placeholder.
'==============================================================================

Public Sub SaveAtt()
    Dim msg As MailItem, exp As Explorer, att As Attachment
    Dim mailIndex As Integer, path As String, folder As String
    Dim Suj As String

    Set exp = Application.ActiveExplorer
    
    ' Archive root — must end with a backslash. Change to suit your environment.
    folder = "\\<server>\<share>\photo catalog\"
    
    mailIndex = 0
    
    For Each msg In exp.Selection
        Suj = msg.Subject
    
        ' Extract the shipment identifier (HM) sitting between "Info DEVRPT"
        ' and "IKEA" in the subject line.
        '   Example subject: "Info DEVRPT 12345678 IKEA Quality Report"
        '   InStr finds the start of each marker; Mid carves out what's between.
        '   The "+ 11" / "- 11" account for the length of "Info DEVRPT".
        HM = Trim(Mid(Suj, _
                      InStr(1, Suj, "Info DEVRPT") + 11, _
                      InStr(1, Suj, "IKEA") - InStr(1, Suj, "Info DEVRPT") - 11))
    
        ' Create the per-shipment folder if it doesn't already exist.
        ' Dir() with vbDirectory returns "" when the path isn't found, so this
        ' is idempotent — safe to re-run.
        If Dir(folder & HM & "\", vbDirectory) = "" Then
            MkDir folder & HM & "\"
        End If
    
        ' Skip emails that have no attachments — keeps the folder clean and
        ' avoids the overhead of touching the file system for nothing.
        If msg.Attachments.Count > 0 Then
            mailIndex = mailIndex + 1
            For Each att In msg.Attachments
                path = folder & HM & "\" & att.FileName
                att.SaveAsFile path
            Next
        End If
    Next
    
    MsgBox "Done!", vbOKOnly + vbInformation, "SaveAtt"
End Sub
