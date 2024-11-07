'*****************************************************************************************
'*****************************************************************************************
'**  Author: Mendeley                                                                   **
'**  Last modified: 2016-04-26                                                          **
'**                                                                                     **
'**  Function refreshDocument(Optional openingDocument As Boolean = False) As Boolean   **
'**                                                                                     **
'**  Refresh the citations in this document and update the citation selector combo-box. **
'**  NOTE: It will not execute out of its context.                                      **
'**  NOTE: Check lines with comment "'MabEntwickeltSich" and apply changes              **
'**     to original macro.                                                              **
'*****************************************************************************************
'*****************************************************************************************
' Refresh the citations in this document and update the
' citation selector combo-box
'
' @param openingDocument Set to true if the refresh is being
' called whilst opening a new document or false if refreshing
' an existing already-open document
Function refreshDocument(Optional openingDocument As Boolean = False) As Boolean
    ' Do not try to refresh a "protected" (e.g. read only, because it's from the internet) doc
    If isProtectedViewDocument() Then
        refreshDocument = False
        Exit Function
    End If
    
    If BENCHMARK_MODE Then
        Dim startTime
        Dim benchmarkTime1
        Dim benchmarkTime2
        Dim benchmarkTime3
        Dim benchmarkTime4

        startTime = Timer()
    End If
    
    Call showStatusBarMessage("Mendeley is preparing to format your citations...")

    Dim currentDocumentPath As String
    currentDocumentPath = activeDocumentPath()

    refreshDocument = False
    Call ActiveDocument.Activate
    
    ZoteroUseBookmarks = False
    
    If openingDocument = True Then
        If Not unitTest Then
            Dim ComboBox2 As CommandBarComboBox
            Set ComboBox2 = getCitationStyleComboBox()
            ComboBox2.Text = getStyleNameFromId(getCitationStyleId())
        End If
        ThisDocument.Saved = True
        Exit Function
    End If
    
    If launchMendeleyIfNecessary() <> CONNECTION_CONNECTED Then
        Exit Function
    End If
    
    If Not isDocumentLinkedToCurrentUser Then
        Exit Function
    End If
    
    Dim documentState As MendeleyDataTypes.DocumentStateType
    documentState = startUpdatingDocument(ActiveDocument)
    ' Update document
    Call beginUndoTransaction("Format Mendeley Citations and Bibliography")
    
    Call sendWordProcessorVersion
    
    Call setCitationStyle(getCitationStyleId())
    If Not unitTest Then
        Call updateCitationStylesComboBox
    End If
    
#If Win32 Then
    If USE_RIBBON Then
        Call recoverRibbonUi
        Call RefreshRibbon
    End If
#End If

    ' Subscribe to events (e.g. WindowSelectionChange) doing on refreshDocument as it
    ' doesn't work in initialise() when addExternalFunctions() is also called
    If Not openingDocument Then
        Set theAppEventHandler.App = Word.Application
    End If

    Dim citationNumberCount As Long
    citationNumberCount = 0
    
    Dim bibliography As String

    Call mendeleyApiClient().resetCitations
    
    Dim marks
    Set marks = fnGetMarks(ZoteroUseBookmarks)
    
    Dim markName As String
    
    Dim thisField As field

    Dim mark

    Dim citationNumber As Long
    citationNumber = 0
    
    For Each mark In marks
        If citationNumber Mod 25 = 0 Then
            Call showStatusBarMessage("Mendeley is preparing to format your citations... (" & _
                Round(100 * citationNumber / marks.count) & "%)")
        End If

        Set thisField = mark
        
        markName = getMarkName(thisField)
        
        If startsWith(markName, "ref Mendeley") Then
            markName = Right(markName, Len(markName) - 4)
            thisField.code.Text = markName
        End If
        
        If isMendeleyCitationField(markName) Then
            citationNumber = citationNumber + 1
            
            ' Just send an empty string if the displayed text is a temporary placeholder
            Dim displayedText As String
            displayedText = getMarkText(thisField)
            'displayedText = getMarkTextWithFormattingTags(thisField)
            If displayedText = INSERT_CITATION_TEXT Or displayedText = MERGING_TEXT Then
                displayedText = ""
            End If
            mendeleyApiClient().addCitation markName, displayedText
            
            thisField.Locked = True
        End If
    Next
    
    Dim oldCitationStyle As String
    oldCitationStyle = getCitationStyleId()
    
    If BENCHMARK_MODE Then
        benchmarkTime1 = Timer() - startTime
        startTime = Timer()
    End If
    
    Call showStatusBarMessage("Mendeley is formatting your citations...")
    
    Call storeTargetDocument

    ' Now that we've compiled the list of uuids, give them to Mendeley Desktop
    ' and tell it to format the citations and bibliography
    If Not mendeleyApiClient().formatCitationsAndBibliography() Then
        Call bringWordToForeground
        Dim errorCitationIndex As Long
        errorCitationIndex = mendeleyApiClient().lastErrorCitationIndex()
        If errorCitationIndex <> -1 Then
            Dim errorField As field
            Dim citationIndex As Long
            citationIndex = mendeleyApiClient().lastErrorCitationIndex() + 1
            Set errorField = marks(citationIndex)
            Call promptToRemoveField(errorField, citationIndex)
        End If
        GoTo ExitFunction
    End If

    Call activateTargetDocument
    
    If BENCHMARK_MODE Then
        benchmarkTime2 = Timer() - startTime
        startTime = Timer()
    End If
    
    citationNumber = 0
    
    Set marks = fnGetMarks(ZoteroUseBookmarks)
    For Each mark In marks 'ActiveDocument.Fields
        If currentDocumentPath <> activeDocumentPath() Then
            GoTo ExitFunction
        End If
        
        If citationNumber Mod 15 = 0 Then
            Call showStatusBarMessage("Mendeley is updating your citations... (" & _
                Round(100 * citationNumber / marks.count) & "%)")
        End If

        Set thisField = mark
        Dim fieldText As String
        Dim plainTextCitation As String
        
        fieldText = ""
        markName = getMarkName(thisField)

        If IsObjectValid(thisField) = False Then
            GoTo NextIterationLoop
        End If
        
        If (isMendeleyCitationField(markName)) Then
            Dim jsonData As String
            jsonData = mendeleyApiClient().getFieldCode(citationNumber)

            If jsonData = "invalid_index" Then
                MsgBox "Mendeley encountered a problem formatting your citations. " & vbCrLf & vbCrLf & _
                    "Please close all other open Word documents and try again."
                GoTo ExitFunction
            End If

            fieldText = mendeleyApiClient().getFormattedCitation(citationNumber)
            plainTextCitation = mendeleyApiClient().getPlainTextFormattedCitation(citationNumber)
            
            Dim previousFormattedCitation As String
            previousFormattedCitation = mendeleyApiClient().getPreviouslyFormattedCitation(citationNumber)

            If currentDocumentPath <> activeDocumentPath() Then
                GoTo ExitFunction
            End If
            Set thisField = fnRenameMark(thisField, jsonData)
            
            If fieldText <> previousFormattedCitation Or plainTextCitation <> getMarkText(thisField) Then
                If currentDocumentPath <> activeDocumentPath() Then
                    GoTo ExitFunction
                End If
                
                ' if Mendeley sends us an empty field, leave it alone since we want to
                ' preserve the user's formatting options
                If fieldText <> "" Then
                    Call applyFormatting(fieldText, thisField)
                End If
            End If
            
            citationNumber = citationNumber + 1
        ElseIf isMendeleyBibliographyField(markName) Then
            If Not InStr(markName, CSL_BIBLIOGRAPHY) > 0 Then
                    Call fnRenameMark(mark, markName & " " & CSL_BIBLIOGRAPHY)
            End If
        
            If bibliography = "" Then
                bibliography = bibliography + mendeleyApiClient().getFormattedBibliography()
                #If Mac Then
                    bibliography = posixToVBAPath(bibliography)
                #End If
            End If
            
            Dim range As range
            Set range = thisField.result
            
            ' Get font used at start of bibliography
            range.Collapse (wdCollapseStart)
            
            Dim currentFontName As String
            Dim currentSize As Long
            currentFontName = range.Font.name
            currentSize = range.Font.size

            ' Get paragraph used at start of bibliography
            Dim currentParagraphStyle As String 'MabEntwickeltSich
            Dim currentParagraphSpaceBefore As Long
            Dim currentParagraphSpaceAfter As Long
            Dim currentLineSpacingRule As Variant

            currentParagraphStyle = range.style 'MabEntwickeltSich
            currentParagraphSpaceBefore = range.ParagraphFormat.SpaceBefore
            currentParagraphSpaceAfter = range.ParagraphFormat.SpaceAfter
            currentLineSpacingRule = range.ParagraphFormat.LineSpacingRule

            ' Insert updated bibliography
            Set range = thisField.result
            ' Word 2013 dirty hack: We can not insert on the whole selection, we need to keep
            ' one character at the end of the selection
            If isWordRangeHackRequired Then
                range.End = range.End - 1
            End If
            range.InsertFile (bibliography)
            ' Word 2013 dirty hack: Remove the last character
            If isWordRangeHackRequired Then
                Set range = thisField.result
                range.Start = range.End - 1
                range.Text = ""
            End If
            
            ' Disable spell and grammar checking on the bibliography.
            ' This is done when the field is created in fnAddMark(), but the InsertFile() call
            ' resets this property of the Field result's range (at least on Mac Word 2011).
            Set range = thisField.result
            range.LanguageID = wdNoProofing
            
            ' Apply font to whole range
            range.Font.name = currentFontName
            range.Font.size = currentSize

            ' Apply paragraph to whole range
            range.ParagraphFormat.SpaceBefore = currentParagraphSpaceBefore
            range.ParagraphFormat.SpaceAfter = currentParagraphSpaceAfter
            range.ParagraphFormat.LineSpacingRule = currentLineSpacingRule

            range.style = currentParagraphStyle 'MabEntwickeltSich
            
            ' Delete first character (is part of the first new paragraph of the RTF file)
            range.End = range.Start + 1
            range.Text = ""
        End If
        
        If Not (fieldText = "") Then
            ' Put text in field
                If thisField.ShowCodes Then
                    thisField.ShowCodes = False
                End If
        End If
        
        thisField.Locked = True
NextIterationLoop:
    Next
    
    If BENCHMARK_MODE Then
        benchmarkTime3 = Timer() - startTime
        startTime = Timer()
    End If
    
    If Not unitTest Then
        Dim newCitationStyle As String
        newCitationStyle = mendeleyApiClient().getCitationStyleId()
        
        If (newCitationStyle <> oldCitationStyle) Then
            ' set new citation style
            Call setCitationStyle(newCitationStyle)
            
            ' update citation styles list
            Call updateCitationStylesComboBox
        End If
        
            Set previouslySelectedField = getFieldAtSelection()
        If Not IsNull(previouslySelectedField) And Not IsEmpty(previouslySelectedField) Then
            previouslySelectedFieldResultText = getMarkText(previouslySelectedField)
        Else
            previouslySelectedFieldResultText = ""
        End If
    End If

    refreshDocument = True

ExitFunction:
    Call endUndoTransaction
    Call finishUpdatingDocument(documentState)
    
    If BENCHMARK_MODE Then
        benchmarkTime4 = Timer() - startTime
    
        MsgBox "Refresh document timings: " & vbCrLf & _
            "pre MD: " & benchmarkTime1 & vbCrLf & _
            "MD: " & benchmarkTime2 & vbCrLf & _
            "post MD: " & benchmarkTime3 & vbCrLf & _
            "post update: " & benchmarkTime4 & vbCrLf & _
            "Total refresh time: " & (benchmarkTime1 + benchmarkTime2 + benchmarkTime3 + benchmarkTime4)
    End If

    Call showStatusBarMessage("")
End Function
