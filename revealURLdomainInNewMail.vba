Sub revealURLdomainInNewMail(ByVal item As Outlook.MailItem)

    idString = "id='" + Environ$("Username") + "' " ' Perhaps using CLASS would be better, but more complicated.
    
    Set myItem = item
    curPos = 1
    msgLen = Len(myItem.HTMLBody)
    Do While (curPos < msgLen)
        nextAnchorStartOpen = InStr(curPos, myItem.HTMLBody, "<a", vbTextCompare)
        If (nextAnchorStartOpen = 0) Then
            Exit Do ' No more anchors.
        Else
            curPos = nextAnchorStartOpen
        End If
        nextAnchorEnd = InStr(curPos, myItem.HTMLBody, "</a>", vbTextCompare)
        If nextAnchorEnd < curPos Then
            Exit Do ' Malformed email.
        End If
        anchor = Mid(myItem.HTMLBody, nextAnchorStartOpen, nextAnchorEnd - nextAnchorStartOpen)
        ' Continue processing anchor if it's not an email address, and it hasn't already been processed (ie idstring already exists in the anchor).
        If (InStr(1, anchor, "href=""mailto:", vbTextCompare) = 0) And (InStr(1, anchor, idString, vbTextCompare) = 0) Then
            ' Get the domain of the link.
            startDomain = InStr(1, anchor, "://", vbTextCompare)
            If startDomain = 0 Then
                GoTo ContinueMessage ' Unsupported Anchor Type
            End If
            anchorStartClose = InStr(startDomain, anchor, ">", vbTextCompare)
            If anchorStartClose = 0 Then
                GoTo ContinueMessage ' Unsupported Anchor Type or malformed anchor
            End If
            endDomain = InStr(startDomain + 3, anchor, "/", vbTextCompare)
            If endDomain = 0 Then
                ' Url may have no URI
                endDomain = InStr(1, anchor, ">", vbTextCompare)
                If endDomain = 0 Then
                    GoTo ContinueMessage ' Malformed anchor
                End If
            End If
            domain = Mid(anchor, startDomain + 3, endDomain - startDomain - 3)
            ' Now that we have the anchor disected, insert the data into the HTMLBody
            
            ' HTMLBody up to the space before the 'href' in the anchor
            leftSide = Left(myItem.HTMLBody, nextAnchorStartOpen + 2)
            'MsgBox "Left (tail):" + vbCrLf + Right(leftSide, 100) + "XXXXXXXX", vbExclamation, "Warning: Skipping Link"
            ' HTMLBody inclusively between the 'href' and the close of the anchor opening.
            middle = Mid(myItem.HTMLBody, nextAnchorStartOpen + 3, anchorStartClose - 3)
            ' HTMLBody from the beginning of the displayed anchor text/data (if image) to the end of the message.
            rightSide = Right(myItem.HTMLBody, msgLen - (nextAnchorStartOpen + anchorStartClose - 1))
            myItem.HTMLBody = leftSide + " " + idString + middle + "[" + domain + "]" + rightSide
            myItem.Save
        End If
ContinueMessage:
        msgLen = Len(myItem.HTMLBody)
        curPos = nextAnchorEnd + 4
    Loop
    Set myItem = Nothing
    
End Sub