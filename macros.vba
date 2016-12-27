'*****************************************************************************************
'*****************************************************************************************
'**  Author: José Luis González García                                                  **
'**  Last modified: 2016-11-15                                                          **
'**                                                                                     **
'**  Sub GAUG_createHyperlinksForCitationsAPA()                                         **
'**                                                                                     **
'**  Generates the bookmarks in the bibliography inserted by Mendeley's plugin.         **
'**  Links the citations inserted by Mendeley's plugin to the corresponding entry       **
'**     in the bibliography inserted by Mendeley's plugin.                              **
'**  Only for APA CSL citation style.                                                   **
'*****************************************************************************************
'*****************************************************************************************
Sub GAUG_createHyperlinksForCitationsAPA()

    Dim documentSection As Section
    Dim sectionField As Field
    Dim blnFound, blnReferenceEntryFound, blnCitationEntryFound, blnCitationEntryPositionFound, blnEditorsFound As Boolean
    Dim intRefereceNumber, intCitationEntryPosition, i As Integer
    Dim objRegExpBiblio, objRegExpCitation, objRegExpCitationData, objRegExpBiblioEntry As RegExp
    Dim colMatchesBiblio, colMatchesCitation, colMatchesCitationData, colMatchesBiblioEntry As MatchCollection
    Dim objMatchBiblio, objMatchCitation, objMatchCitationData As Match
    Dim strTempMatch As String


'*****************************
'*  Cleaning old hyperlinks  *
'*****************************
    GAUG_removeHyperlinksForCitations






'*****************************
'*   Creation of bookmarks   *
'*****************************
    'creates the object for regular expressions (to get all entries in biblio)
    Set objRegExpBiblio = New RegExp
    'sets the pattern to match every reference in bibliography (it may include character of carriage return)
    '(all text from the beginning of the string or carriage return until a year between parentheses is found)
    objRegExpBiblio.Pattern = "((^)|(\r))[^(\r)]*\(\d\d\d\d[a-zA-Z]?\)"
    'sets case insensitivity
    objRegExpBiblio.IgnoreCase = False
    'sets global applicability
    objRegExpBiblio.Global = True

    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        'checks if the section has text with style "Titre de dernière section"
        '(it is a section not belonging to any chapter after the sections of parts and chapters)
        blnFound = False
        With documentSection.Range.Find
            .Style = "Titre de dernière section"
            .Execute
            blnFound = .Found
        End With

        'checks if the bibliography is in this section
        If blnFound Then
            'checks all fields
            For Each sectionField In documentSection.Range.Fields
                'if it is the bibliography
                If sectionField.Type = wdFieldAddin And Trim(sectionField.Code) = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY" Then
                    'start the numbering
                    intRefereceNumber = 1

                    'selects the current field (Mendeley's bibliography field)
                    sectionField.Select

                    'checks that the string can be compared
                    If (objRegExpBiblio.Test(Selection) = True) Then
                        'gets the matches (all entries in bibliography according to the regular expression)
                        Set colMatchesBiblio = objRegExpBiblio.Execute(Selection)

                        'treats all matches (all entries in bibliography) to generate bookmars
                        '(we have to find AGAIN every entry to select it and create the bookmark)
                        For Each objMatchBiblio In colMatchesBiblio
                            'removes the carriage return from match, if necessary
                            strTempMatch = Replace(objMatchBiblio.Value, Chr(13), "")

                            'selects the current field (Mendeley's bibliography field)
                            sectionField.Select

                            'finds and selects the text of the current reference
                            blnReferenceEntryFound = False
                            With Selection.Find
                                .Forward = True
                                .Wrap = wdFindStop
                                .Text = strTempMatch
                                .Execute
                                blnReferenceEntryFound = .Found
                            End With

                            'if a match was found (it shall always find it, but good practice)
                            'creates the bookmark with the selected text
                            If blnReferenceEntryFound Then
                                'creates the bookmark
                                Selection.Bookmarks.Add _
                                    Name:="SignetBibliographie_" & Format(CStr(intRefereceNumber), "00#"), _
                                    Range:=Selection.Range
                            End If

                            'continues with the next number
                            intRefereceNumber = intRefereceNumber + 1

                        Next
                    End If
                    'by now, we have created all bookmarks and have all entries in colMatchesBiblio
                    'for future use when creating the hyperlinks

                End If 'if it is the biblio
            Next 'sectionField
        End If
    Next 'documentSection






'*****************************
'*   Linking the bookmarks   *
'*****************************
    'creates the object for regular expressions (to get all entries in current citation, all entries of data in current citation, position of citation entry in biblio)
    Set objRegExpCitation = New RegExp
    Set objRegExpCitationData = New RegExp
    Set objRegExpBiblioEntry = New RegExp
    'sets the pattern to match every citation entry (with or without autors) in current field
    '(the year of publication is always present, authors may not be present)
    '(all text non starting by "(" or "," or ";" followed by non digits until a year is found)
    objRegExpCitation.Pattern = "[^(\(|,|;)][^0-9]*\d\d\d\d[a-zA-Z]?"
    'sets the pattern to match every citation entry from the data of the current field
    'original regular expression to get the authors info from Field.Code "((\"id\")|(\"family\")|(\"given\"))\s\:\s\"[^\"]*\""
    '(all text related to "id", "family" and "given")
    'objRegExpCitationData.Pattern = "((\" & Chr(34) & "id\" & Chr(34) & ")|(\" & Chr(34) & "family\" & Chr(34) & ")|(\" & Chr(34) & "given\" & Chr(34) & "))\s\:\s\" & Chr(34) & "[^\" & Chr(34) & "]*\" & Chr(34)
    'updated to separate authors from editors:
    objRegExpCitationData.Pattern = "(\" & Chr(34) & "editor\" & Chr(34) & "\s:\s)|(((\" & Chr(34) & "id\" & Chr(34) & ")|(\" & Chr(34) & "family\" & Chr(34) & "))\s\:\s\" & Chr(34) & "[^\" & Chr(34) & "]*\" & Chr(34) & ")"
    'sets case insensitivity
    objRegExpCitation.IgnoreCase = False
    objRegExpCitationData.IgnoreCase = False
    objRegExpBiblioEntry.IgnoreCase = False
    'sets global applicability
    objRegExpCitation.Global = True
    objRegExpCitationData.Global = True
    objRegExpBiblioEntry.Global = True

    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        'checks all fields
        For Each sectionField In documentSection.Range.Fields
            'if it is a citation
            If sectionField.Type = wdFieldAddin And Left(sectionField.Code, 18) = "ADDIN CSL_CITATION" Then

                'selects the current field (Mendeley's citation field)
                sectionField.Select

                'checks that the string can be compared (both, Selection and Field.Code)
                If (objRegExpCitation.Test(Selection) = True) And (objRegExpCitationData.Test(sectionField.Code) = True) Then
                    'gets the matches (all entries in the citation according to the regular expression)
                    Set colMatchesCitation = objRegExpCitation.Execute(Selection)
                    'gets the matches (all entries in the citation .Data according to the regular expression)
                    '(used to find the entry in the bibliography)
                    Set colMatchesCitationData = objRegExpCitationData.Execute(sectionField.Code)

                    'resets the position (current citation entry being treated)
                    intCitationEntryPosition = 1

                    'treats all matches (all entries in citation) to generate hyperlinks
                    For Each objMatchCitation In colMatchesCitation
                        'I COULD NOT FIND A MORE EFFICIENT WAY TO SELECT EVERY REFERENCE
                        'IN ORDER TO CREATE THE LINK:
                        'Start: Needs re-work

                        'flag to find the position of the current citation entry
                        blnCitationEntryPositionFound = False

                        'flag to skip the name of the editors
                        blnEditorsFound = False

                        'gets the data from current citation entry to build the pattern to find the reference entry in biblio
                        objRegExpBiblioEntry.Pattern = ""
                        For Each objMatchCitationData In colMatchesCitationData
                            'activates the flag only if in current citation entry
                            'if the current citation entry starts/ends here ("id" : "ITEM-X")
                            If objMatchCitationData.Value = Chr(34) & "id" & Chr(34) & " : " & Chr(34) & "ITEM-" & CStr(intCitationEntryPosition) & Chr(34) Then
                                blnCitationEntryPositionFound = Not blnCitationEntryPositionFound
                            Else
                                If blnCitationEntryPositionFound Then
                                    'if the "editor" names start here
                                    If objMatchCitationData.Value = Chr(34) & "editor" & Chr(34) & " : " Then
                                        blnEditorsFound = True
                                    Else
                                        'if the names are the author's names
                                        If Not blnEditorsFound Then
                                            'gets the last name of the author and adds it to the regular expression
                                            objRegExpBiblioEntry.Pattern = objRegExpBiblioEntry.Pattern & Replace(Mid(objMatchCitationData.Value, InStr(objMatchCitationData.Value, Chr(34) & " : " & Chr(34)) + 5), Chr(34), "") & ".*"
                                        End If
                                    End If
                                End If
                            End If

                        Next

                        'adds the year of current citation entry
                        If Mid(objMatchCitation.Value, Len(objMatchCitation.Value) - 4, 1) = " " Then
                            objRegExpBiblioEntry.Pattern = objRegExpBiblioEntry.Pattern & "\(" & Right(objMatchCitation.Value, 4) & "\)"
                        Else
                            objRegExpBiblioEntry.Pattern = objRegExpBiblioEntry.Pattern & "\(" & Right(objMatchCitation.Value, 5) & "\)"
                        End If

                        'at this point, the regular expression to find the entry in the biblio is ready

                        'initializes the position
                        i = 1
                        'finds the position of the citation entry in the list of references in the biblio
                        blnReferenceEntryFound = False
                        For Each objMatchBiblio In colMatchesBiblio
                            'gets the matches, if any, to check if this reference entry corresponds to the citation being treated
                            Set colMatchesBiblioEntry = objRegExpBiblioEntry.Execute(objMatchBiblio.Value)
                            'if the this is the corresponding reference entry
                            'Verify for MabEntwickeltSich: perhaps a more strict verification is needed
                            If colMatchesBiblioEntry.Count > 0 Then
                                blnReferenceEntryFound = True
                                Exit For
                            End If
                            'continues with the next number
                            i = i + 1
                        Next

                        'at this point we also have the position (i) in the biblio, we are ready to create the hyperlink

                        'if reference entry was found (shall always find it), creates the hyperlink
                        If blnReferenceEntryFound Then
                            'selects the current field (Mendeley's citation field)
                            sectionField.Select

                            'finds the opening parenthesis (first character of the field),
                            'used to select something inside the field
                            With Selection.Find
                                .Forward = True
                                .Wrap = wdFindStop
                                .Text = "("
                                .Execute
                                blnCitationEntryFound = .Found
                            End With

                            'if a match was found (it should always find it, but good practice)
                            'selects the correct entry text from the citation field
                            If blnCitationEntryFound Then
                                'recalculates the selection to include only the match (entry) in citation
                                Selection.MoveEnd Unit:=wdCharacter, Count:=objMatchCitation.FirstIndex + objMatchCitation.Length - 1
                                'if the first character is a blank space
                                If Left(objMatchCitation.Value, 1) = " " Then
                                    'removes the leading blank space
                                    Selection.MoveStart Unit:=wdCharacter, Count:=objMatchCitation.FirstIndex + 1
                                Else
                                    'uses the whole range
                                    Selection.MoveStart Unit:=wdCharacter, Count:=objMatchCitation.FirstIndex
                                End If

                                'creates the hyperlink for the current citation entry
                                'a cross-reference is not a good idea, it changes the text in citation (or may delete citation):
                                'Selection.Fields.Add Range:=Selection.Range, _
                                '    Type:=wdFieldEmpty, _
                                '    Text:="REF " & Chr(34) & "SignetBibliographie_" & Format(CStr(i), "00#") & Chr(34) & " \h", _
                                '    PreserveFormatting:=True
                                'better to use normal hyperlink:
                                Selection.Hyperlinks.Add Anchor:=Selection.Range, _
                                    Address:="", SubAddress:="SignetBibliographie_" & Format(CStr(i), "00#"), _
                                    ScreenTip:=""

                            End If
                        Else
                            MsgBox ("Orphan citation entry found:" & vbCrLf & vbCrLf & objMatchCitation.Value & vbCrLf & vbCrLf & "Remove it from document!")
                        End If

                        'Ends: Needs re-work

                        'at this point current citation entry is linked to corresponding reference in bilio

                        'continues with the next position (next citation entry)
                        intCitationEntryPosition = intCitationEntryPosition + 1

                    Next 'treats all matches (all entries in citation) to generate hyperlinks

                End If 'checks that the string can be compared

            End If 'if it is a citation
        Next 'sectionField

        'at this point all citations are linked to their corresponding reference in bilio

    Next 'documentSection

End Sub



'*****************************************************************************************
'*****************************************************************************************
'**  Author: José Luis González García                                                  **
'**  Last modified: 2016-11-15                                                          **
'**                                                                                     **
'**  Sub GAUG_createHyperlinksForCitationsIEEE()                                        **
'**                                                                                     **
'**  Generates the bookmarks in the bibliography inserted by Mendeley's plugin.         **
'**  Links the citations inserted by Mendeley's plugin to the corresponding entry       **
'**     in the bibliography inserted by Mendeley's plugin.                              **
'**  Only for IEEE CSL citation style.                                                  **
'*****************************************************************************************
'*****************************************************************************************
Sub GAUG_createHyperlinksForCitationsIEEE()

    Dim documentSection As Section
    Dim sectionField As Field
    Dim blnFound, blnReferenceNumberFound, blnCitationNumberFound As Boolean
    Dim intRefereceNumber As Integer


'*****************************
'*  Cleaning old hyperlinks  *
'*****************************
    GAUG_removeHyperlinksForCitations






'*****************************
'*   Creation of bookmarks   *
'*****************************
    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        'checks if the section has text with style "Titre de dernière section"
        '(it is a section not belonging to any chapter after the sections of parts and chapters)
        blnFound = False
        With documentSection.Range.Find
            .Style = "Titre de dernière section"
            .Execute
            blnFound = .Found
        End With

        'checks if the bibliography is in this section
        If blnFound Then
            'checks all fields
            For Each sectionField In documentSection.Range.Fields
                'if it is the bibliography
                If sectionField.Type = wdFieldAddin And Trim(sectionField.Code) = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY" Then
                    'start the numbering
                    intRefereceNumber = 1
                    Do
                        'selects the current field (Mendeley's citation field)
                        sectionField.Select

                        'finds and selects the text of the number of the reference
                        With Selection.Find
                            .Forward = True
                            .Wrap = wdFindStop
                            .Text = "[" & CStr(intRefereceNumber) & "]"
                            .Execute
                            blnReferenceNumberFound = .Found
                        End With

                        'if a number of a reference was found, creates the bookmark with the selected text
                        If blnReferenceNumberFound Then
                            'restricts the selection to only the number
                            With Selection.Find
                                .Forward = True
                                .Wrap = wdFindStop
                                .Text = CStr(intRefereceNumber)
                                .Execute
                                blnReferenceNumberFound = .Found
                            End With

                            'creates the bookmark
                            Selection.Bookmarks.Add _
                                Name:="SignetBibliographie_" & Format(CStr(intRefereceNumber), "00#"), _
                                Range:=Selection.Range
                        End If

                        'continues with the next number
                        intRefereceNumber = intRefereceNumber + 1

                    'while numbers of refereces are found
                    Loop While (blnReferenceNumberFound)
                End If 'if it is the biblio
            Next 'sectionField
        End If
    Next 'documentSection






'*****************************
'*   Linking the bookmarks   *
'*****************************
    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        'checks all fields
        For Each sectionField In documentSection.Range.Fields
            'if it is a citation
            If sectionField.Type = wdFieldAddin And Left(sectionField.Code, 18) = "ADDIN CSL_CITATION" Then

                'check for all numbers of citations
                For i = 1 To intRefereceNumber

                    'selects the current field (Mendeley's citation field)
                    sectionField.Select

                    'finds and selects the text of the number of the citation
                    With Selection.Find
                        .Forward = True
                        .Wrap = wdFindStop
                        .Text = "[" & CStr(i) & "]"
                        .Execute
                        blnCitationNumberFound = .Found
                    End With

                    'if a number of a citation was found, inserts the hyperlink
                    If blnCitationNumberFound Then
                        'restricts the selection to only the number
                        With Selection.Find
                            .Forward = True
                            .Wrap = wdFindStop
                            .Text = CStr(i)
                            .Execute
                            blnCitationNumberFound = .Found
                        End With

                        'a cross-reference is not a good idea, it changes the text in citation (or may delete citation):
                        'Selection.Fields.Add Range:=Selection.Range, _
                        '    Type:=wdFieldEmpty, _
                        '    Text:="REF " & Chr(34) & "SignetBibliographie_" & Format(CStr(i), "00#") & Chr(34) & " \h", _
                        '    PreserveFormatting:=True
                        'better to use normal hyperlink:
                        Selection.Hyperlinks.Add Anchor:=Selection.Range, _
                            Address:="", SubAddress:="SignetBibliographie_" & Format(CStr(i), "00#"), _
                            ScreenTip:=""
                    End If
                Next 'all numbers of citations

            End If 'if it is a citation
        Next 'sectionField
    Next 'documentSection

End Sub



'*****************************************************************************************
'*****************************************************************************************
'**  Author: José Luis González García                                                  **
'**  Last modified: 2016-04-26                                                          **
'**                                                                                     **
'**  Sub GAUG_removeHyperlinksForCitations()                                            **
'**                                                                                     **
'**  Removes the bookmarks generated by GAUG_createHyperlinksForCitations               **
'**     in the bibliography inserted by Mendeley's plugin.                              **
'**  Removes the hyperlinks generated by GAUG_createHyperlinksForCitations              **
'**     of the citations inserted by Mendeley's plugin.                                 **
'*****************************************************************************************
'*****************************************************************************************
Sub GAUG_removeHyperlinksForCitations()
    Dim documentSection As Section
    Dim sectionField As Field
    Dim blnFound As Boolean



'*****************************
'*  Cleaning old hyperlinks  *
'*****************************
    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        For Each sectionField In documentSection.Range.Fields
            'if it is a citation
            If sectionField.Type = wdFieldAddin And Left(sectionField.Code, 18) = "ADDIN CSL_CITATION" Then
                'tries to restore it to its original value, first with the whole field
                sectionField.Select
                GAUG_getUndoEditButton().Execute
            End If
        Next



        'checks if the section has text with style "Titre de dernière section"
        '(it is a section not belonging to any chapter after the sections of parts and chapters)
        blnFound = False
        With documentSection.Range.Find
            .Style = "Titre de dernière section"
            .Execute
            blnFound = .Found
        End With

        'checks if the bibliography is in this section
        If blnFound Then
            For Each sectionField In documentSection.Range.Fields
                'if it is the bibliography
                If sectionField.Type = wdFieldAddin And Trim(sectionField.Code) = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY" Then
                    'deletes old bookmarks
                    sectionField.Select
                    For Each fieldBookmark In Selection.Bookmarks
                        'deletes current bookmark
                        fieldBookmark.Delete
                    Next
                End If
            Next
        End If

    Next

End Sub



'*****************************************************************************************
'*****************************************************************************************
'**  Author: Mendeley                                                                   **
'**  Last modified: 2016-04-26                                                          **
'**                                                                                     **
'**  Function GAUG_getUndoEditButton() As CommandBarButton                              **
'**                                                                                     **
'**  Gets the CommandBarButton "Undo Edit" installed by Mendeley's plugin.              **
'**  The CommandBarButton is used to restore the original citation fields               **
'**     inserted by Mendeley.                                                           **
'*****************************************************************************************
'*****************************************************************************************
Function GAUG_getUndoEditButton() As CommandBarButton 'copied from Mendeley's plugin function "getUndoEditButton"

    Dim mendeleyControl As CommandBarControl

    For Each mendeleyControl In CommandBars("Mendeley Toolbar").Controls
        If mendeleyControl.Caption = "Undo Edit" Then
            Set GAUG_getUndoEditButton = mendeleyControl
            Exit Function
        End If
    Next
    ' if here, button hasn't been created yet
    MsgBox "Undo edit button not found"
End Function



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
