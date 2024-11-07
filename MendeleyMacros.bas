Attribute VB_Name = "MendeleyMacros"
'*****************************************************************************************
'*****************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                  **
'**  Last modified: 2020-07-25                                                          **
'**                                                                                     **
'**  Sub GAUG_createHyperlinksForCitationsAPA()                                         **
'**                                                                                     **
'**  Generates the bookmarks in the bibliography inserted by Mendeley's plugin.         **
'**  Links the citations inserted by Mendeley's plugin to the corresponding entry       **
'**     in the bibliography inserted by Mendeley's plugin.                              **
'**  Generates the hyperlinks for the URLs in the bibliography inserted by              **
'**     Mendeley's plugin.                                                              **
'**  Only for APA CSL citation style.                                                   **
'*****************************************************************************************
'*****************************************************************************************
Sub GAUG_createHyperlinksForCitationsAPA()

    Dim documentSection As Section
    Dim sectionField As Field
    Dim blnFound, blnBibliographyFound, blnReferenceEntryFound, blnCitationEntryFound, blnCitationEntryPositionFound, blnEditorsFound, blnAuthorsFound, blnGenerateHyperlinksForURLs, blnURLFound As Boolean
    Dim intRefereceNumber, intCitationEntryPosition, i As Integer
    Dim objRegExpBiblio, objRegExpCitation, objRegExpCitationData, objRegExpBiblioEntry, objRegExpFindEntry, objRegExpURL As RegExp
    Dim colMatchesBiblio, colMatchesCitation, colMatchesCitationData, colMatchesBiblioEntry, colMatchesFindEntry, colMatchesURL As MatchCollection
    Dim objMatchBiblio, objMatchCitation, objMatchCitationData, objMatchFindEntry, objMatchURL As match
    Dim strTempMatch, strSubStringOfTempMatch, strLastAuthors, strLastYear As String
    Dim strTypeOfExecution As String
    Dim blnMabEntwickeltSich As Boolean
    Dim stlStyleInDocument As Word.Style
    Dim strStyleForTitleOfBibliography As String
    Dim blnStyleForTitleOfBibliographyFound As Boolean
    Dim strURL As String
    Dim arrNonDetectedURLs, varNonDetectedURL As Variant


'*****************************
'*   Custom configuration    *
'*****************************
    'SEE DOCUMENTATION
    'specifies the name of the font style used for the title of the bibliography
    strStyleForTitleOfBibliography = "Titre de derni�re section"

    'possible values are True and False
    'SEE DOCUMENTATION
    'set to True if the bibliography is in a section with title in style indicated by strStyleForTitleOfBibliography
    blnMabEntwickeltSich = False

    'possible values are True and False
    'SEE DOCUMENTATION
    'if set to True, then the URLs in the bibliography will be converted to hyperlinks
    blnGenerateHyperlinksForURLs = True

    'SEE DOCUMENTATION
    'specifies the URLs, not detected by the regular expression, to generate the hyperlinks in the bibliography
    'note that the last URL does not have a comma after the double quotation mark
    arrNonDetectedURLs = _
        Array( _
            "http://my.first.sample/url/", _
            "https://my.second.sample/url/", _
            "ftp://my.third.sample/url/with/file.pdf" _
            )

    'possible values are "RemoveHyperlinks", "CleanEnvironment" and "CleanFullEnvironment"
    'SEE DOCUMENTATION
    strTypeOfExecution = "RemoveHyperlinks"






'*****************************
'*     Initial tasks and     *
'*       verifications       *
'*****************************
    'checks that the style exists in the collection of styles of the document
    For Each stlStyleInDocument In ActiveDocument.Styles
        'checks if the current style is the style for the title of the bibliography
        If stlStyleInDocument = strStyleForTitleOfBibliography Then
            blnStyleForTitleOfBibliographyFound = True
        End If
    Next 'all styles in document

    'if the style was not found
    If blnMabEntwickeltSich And Not blnStyleForTitleOfBibliographyFound Then
        MsgBox "The style " & Chr(34) & strStyleForTitleOfBibliography & Chr(34) & " could not be found." & vbCrLf & vbCrLf & _
            "Add the custom style to Microsoft Word or" & vbCrLf & _
            "set the flag blnMabEntwickeltSich to False." & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_createHyperlinksForCitationsAPA()"

        'stops the execution
        End
    End If






'*****************************
'*  Cleaning old hyperlinks  *
'*****************************
    'removes all hyperlinks
    Call GAUG_removeHyperlinksForCitations(strTypeOfExecution)






    'disables the screen updating while creating the hyperlinks to increase speed
    Application.ScreenUpdating = False

'*****************************
'*   Creation of bookmarks   *
'*****************************
    'creates the object for regular expressions (to get all entries in biblio)
    Set objRegExpBiblio = New RegExp
    'sets the pattern to match every reference in bibliography (it may include character of carriage return)
    '(all text from the beginning of the string or carriage return until a year between parentheses is found)
    'objRegExpBiblio.Pattern = "((^)|(\r))[^(\r)]*\(\d\d\d\d[a-zA-Z]?\)"
    'updated to include "(Ed.)" and "(Eds.)" when editors are used for the citations and bibliography
    objRegExpBiblio.Pattern = "((^)|(\r))[^(\r)]*(\(Eds?\.\)\.\s)?\(\d\d\d\d[a-zA-Z]?\)"
    'sets case insensitivity
    objRegExpBiblio.IgnoreCase = False
    'sets global applicability
    objRegExpBiblio.Global = True

    'creates the object for regular expressions (to get all URLs in biblio
    Set objRegExpURL = New RegExp
    'sets the pattern to match every URL in bibliography
    objRegExpURL.Pattern = "((https?)|(ftp)):\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z0-9]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=\[\]\(\)<>;]*)"
    'sets case insensitivity
    objRegExpURL.IgnoreCase = False
    'sets global applicability
    objRegExpURL.Global = True

    'initializes the flag
    blnBibliographyFound = False
    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        'if this is MabEntwickeltSich's document structure
        If blnMabEntwickeltSich Then
            'checks if the section has text with style indicated by strStyleForTitleOfBibliography
            '(it is a section not belonging to any chapter after the sections of parts and chapters)
            blnFound = False
            With documentSection.range.Find
                .Style = strStyleForTitleOfBibliography
                .Execute
                blnFound = .found
            End With
        'this is a document by another user
        Else
            'forces the macro to search for the bibliography in every section
            blnFound = True
        End If

        'checks if the bibliography is in this section
        If blnFound Then
            'checks all fields
            For Each sectionField In documentSection.range.Fields
                'if it is the bibliography
                If sectionField.Type = wdFieldAddin And Trim(sectionField.Code) = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY" Then
                    blnBibliographyFound = True
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
                            strTempMatch = Replace(objMatchBiblio.value, Chr(13), "")

                            'prevents errors if the match is longer than 256 characters
                            'however, the reference will not be linked if it has more than 20 authors (APA 7th edition)
                            'or more than 7 authors (APA 6th edition) due to the fact that APA replaces some of them with ellipsis (...)
                            strSubStringOfTempMatch = Left(strTempMatch, 256)

                            'selects the current field (Mendeley's bibliography field)
                            sectionField.Select

                            'finds and selects the text of the current reference
                            blnReferenceEntryFound = False
                            With Selection.Find
                                .Forward = True
                                .Wrap = wdFindStop
                                .Text = strSubStringOfTempMatch
                                .Execute
                                blnReferenceEntryFound = .found
                            End With

                            'moves the selection, if necessary, to include the full match
                            Selection.MoveEnd Unit:=wdCharacter, Count:=Len(strTempMatch) - Len(strSubStringOfTempMatch)

                            'checks that the full match is found
                            If Selection.Text = strTempMatch Then
                                blnReferenceEntryFound = True
                            Else
                                'there is no more searching, the reference will not be linked in this case
                                blnReferenceEntryFound = False
                            End If

                            'if a match was found (it shall always find it, but good practice)
                            'creates the bookmark with the selected text
                            If blnReferenceEntryFound Then
                                'creates the bookmark
                                Selection.Bookmarks.Add _
                                    Name:="SignetBibliographie_" & format(CStr(intRefereceNumber), "00#"), _
                                    range:=Selection.range
                            End If

                            'continues with the next number
                            intRefereceNumber = intRefereceNumber + 1

                        Next
                    End If
                    'by now, we have created all bookmarks and have all entries in colMatchesBiblio
                    'for future use when creating the hyperlinks

                    'generates the hyperlinks for the URLs in the bibliography if required
                    If blnGenerateHyperlinksForURLs Then

                        'generates the hyperlnks from the list of non detected URLs
                        'the non detected URLs shall be done first or some conflicts may arise
                        For Each varNonDetectedURL In arrNonDetectedURLs
                                'selects the current field (Mendeley's bibliography field)
                                sectionField.Select

                                'finds all instances of current URL
                                Do
                                    'finds and selects the text of the URL
                                    With Selection.Find
                                        .Forward = True
                                        .Wrap = wdFindStop
                                        .Text = CStr(varNonDetectedURL)
                                        .Execute
                                        blnURLFound = .found
                                    End With

                                    'creates the hyperlink
                                    If blnURLFound Then
                                        'checks there is no hyperlink already
                                        If Selection.Hyperlinks.Count = 0 Then
                                            Selection.Hyperlinks.Add Anchor:=Selection.range, _
                                                Address:=Replace(Trim(CStr(varNonDetectedURL)), " ", "%20"), SubAddress:="", _
                                                ScreenTip:=""
                                        End If
                                    End If

                                Loop Until (Not blnURLFound) 'finds all instances of current URL
                        Next 'generates the hyperlnks from the list of non detected URLs

                        'selects the current field (Mendeley's bibliography field)
                        sectionField.Select

                        'checks that the string can be compared (both, Selection and Field.Code)
                        If objRegExpURL.Test(Selection) = True Then
                            'gets the matches (all URLs in the biblio according to the regular expression)
                            Set colMatchesURL = objRegExpURL.Execute(Selection)

                            'treats all matches (all URLs in biblio) to generate hyperlinks
                            For Each objMatchURL In colMatchesURL

                                'removes the last character if a period (".")
                                If Right(objMatchURL.value, 1) = "." Then
                                    strURL = Left(objMatchURL.value, Len(objMatchURL.value) - 1)
                                Else
                                    strURL = objMatchURL.value
                                End If

                                'selects the current field (Mendeley's bibliography field)
                                sectionField.Select

                                'finds all instances of current URL
                                Do
                                    'finds and selects the text of the URL
                                    With Selection.Find
                                        .Forward = True
                                        .Wrap = wdFindStop
                                        .Text = strURL
                                        .Execute
                                        blnURLFound = .found
                                    End With

                                    'creates the hyperlink
                                    If blnURLFound Then
                                        'checks there is no hyperlink already
                                        'this could happen in some cases:
                                        '     when duplicated URLs in bibliography
                                        '     when the hyperlink was created with the list of non detected URLs
                                        '     when a partial URL is detected and the hyperlink was created with the list of non detected URLs
                                        If Selection.Hyperlinks.Count = 0 Then
                                            Selection.Hyperlinks.Add Anchor:=Selection.range, _
                                                Address:=strURL, SubAddress:="", _
                                                ScreenTip:=""
                                        End If
                                    End If

                                Loop Until (Not blnURLFound) 'finds all instances of current URL

                            Next 'treats all matches (all URLs in biblio) to generate hyperlinks
                        End If 'checks that the string can be compared

                    End If 'hyperlinks for URLs in bibliography

                End If 'if it is the biblio
            Next 'sectionField
        End If
    Next 'documentSection


    'if the bibliography could not be located in the document
    If Not blnBibliographyFound Then
        MsgBox "The bibliography could not be located in the document." & vbCrLf & vbCrLf & _
            "Make sure that you have inserted the bibliography via the Mendeley's plugin" & vbCrLf & _
            "and that the custom configuration of the GAUG_* macros is correct." & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_createHyperlinksForCitationsAPA()"

        'stops the execution
        End
    End If






'*****************************
'*   Linking the bookmarks   *
'*****************************
    'creates the object for regular expressions (to get all entries in current citation, all entries of data in current citation, position of citation entry in biblio)
    Set objRegExpCitation = New RegExp
    Set objRegExpCitationData = New RegExp
    Set objRegExpBiblioEntry = New RegExp
    Set objRegExpFindEntry = New RegExp
    'sets the pattern to match every citation entry (with or without authors) in current field
    '(the year of publication is always present, authors may not be present)
    '(all text non starting by "(" or "," or ";" followed by non digits until a year is found)
    objRegExpCitation.Pattern = "[^(\(|,|;)][^0-9]*\d\d\d\d[a-zA-Z]?"
    'sets the pattern to match every citation entry from the data of the current field
    'original regular expression to get the authors info from Field.Code "((\"id\")|(\"family\")|(\"given\"))\s\:\s\"[^\"]*\""
    '(all text related to "id", "family" and "given"), all '\"' were substituted by '\" & Chr(34) & "'
    'objRegExpCitationData.Pattern = "((\" & Chr(34) & "id\" & Chr(34) & ")|(\" & Chr(34) & "family\" & Chr(34) & ")|(\" & Chr(34) & "given\" & Chr(34) & "))\s\:\s\" & Chr(34) & "[^\" & Chr(34) & "]*\" & Chr(34)
    'updated to separate authors from editors:
    'objRegExpCitationData.Pattern = "(\" & Chr(34) & "editor\" & Chr(34) & "\s:\s)|(((\" & Chr(34) & "id\" & Chr(34) & ")|(\" & Chr(34) & "family\" & Chr(34) & "))\s\:\s\" & Chr(34) & "[^\" & Chr(34) & "]*\" & Chr(34) & ")"
    'updated to also include the publication year:
    objRegExpCitationData.Pattern = "((\" & Chr(34) & "editor\" & Chr(34) & "\s*:\s*)|(((\" & Chr(34) & "id\" & Chr(34) & ")|(\" & Chr(34) & "family\" & Chr(34) & "))\s*\:\s*\" & Chr(34) & "[^\" & Chr(34) & "]*\" & Chr(34) & "))|(\[\s*\[\s*\" & Chr(34) & "[0-9]+\" & Chr(34) & "\s*\]\s*\])"
    'sets case insensitivity
    objRegExpCitation.IgnoreCase = False
    objRegExpCitationData.IgnoreCase = False
    objRegExpBiblioEntry.IgnoreCase = False
    objRegExpFindEntry.IgnoreCase = False
    'sets global applicability
    objRegExpCitation.Global = True
    objRegExpCitationData.Global = True
    objRegExpBiblioEntry.Global = True
    objRegExpFindEntry.Global = True

    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        'checks all fields
        For Each sectionField In documentSection.range.Fields
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

                    'treats all matches (all entries in citation) to generate hyperlinks
                    For Each objMatchCitation In colMatchesCitation
                        'I COULD NOT FIND A MORE EFFICIENT WAY TO SELECT EVERY REFERENCE
                        'IN ORDER TO CREATE THE LINK:
                        'Start: Needs re-work

                        'when citations are merged, they are ordered by the authors' family names
                        'the position of the citation in the visible text may not correspond to the position in the citation hidden data,
                        'we need to find the entry, but we may not have the authors's family names :(

                        'if the current match has authors's family names (not only the year)
                        'we keep them stored for future use if needed
                        If Len(Trim(objMatchCitation.value)) > 6 Then 'includes ", " before the year
                            strLastAuthors = objMatchCitation.value
                            'removes the last character that could be a letter, the next loop will finish removing the year
                            strLastAuthors = Left(strLastAuthors, Len(strLastAuthors) - 1)
                        End If
                        'removes the years to leave only the authors's family names
                        Do While IsNumeric(Right(strLastAuthors, 1)) Or (Right(strLastAuthors, 1) = ",") Or (Right(strLastAuthors, 1) = " ")
                            strLastAuthors = Left(strLastAuthors, Len(strLastAuthors) - 1)
                        Loop
                        strLastAuthors = Trim(strLastAuthors) '"et al." may still be in the string, but we need it that way


                        'iterates to find all ("id" : "ITEM-X") in colMatchesCitationData to identify where the citation is located
                        For intCitationEntryPosition = 1 To colMatchesCitation.Count

                            'flag to find the position of the current citation entry
                            blnCitationEntryPositionFound = False

                            'flag to skip the name of the editors
                            blnEditorsFound = False

                            'initializes the regular expressions
                            objRegExpBiblioEntry.Pattern = ""
                            objRegExpFindEntry.Pattern = ""

                            'gets the data from current citation entry to build the pattern to find the reference entry in biblio
                            For Each objMatchCitationData In colMatchesCitationData
                                'activates the flag only if in current citation entry
                                'if the current citation entry starts/ends here ("id" : "ITEM-X")
                                If objMatchCitationData.value = Chr(34) & "id" & Chr(34) & " : " & Chr(34) & "ITEM-" & CStr(intCitationEntryPosition) & Chr(34) Then
                                    blnCitationEntryPositionFound = Not blnCitationEntryPositionFound
                                Else
                                    If blnCitationEntryPositionFound Then
                                        'if the "editor" names start here, sets the flag to stop adding them
                                        If objMatchCitationData.value = Chr(34) & "editor" & Chr(34) & " : " Then
                                            'but if no authors were found (like with a book with only editors), then the flag is not set because the editors are used for the citation
                                            If Len(objRegExpFindEntry.Pattern) > 0 Then
                                                blnEditorsFound = True
                                            End If
                                        Else
                                            'skips the year related to "accessed" that may be between start/end of current ("id" : "ITEM-X")
                                            If Not (Left(objMatchCitationData.value, 5) = "[ [ " & Chr(34) And Right(objMatchCitationData.value, 5) = Chr(34) & " ] ]") Then
                                                'if the names are the author's names
                                                If Not blnEditorsFound Then
                                                    'gets the last name of the author and adds it to the regular expression
                                                    objRegExpBiblioEntry.Pattern = objRegExpBiblioEntry.Pattern & Replace(Mid(objMatchCitationData.value, InStr(objMatchCitationData.value, Chr(34) & " : " & Chr(34)) + 5), Chr(34), "") & ".*"
                                                    'creates another patterns to match the citation entry with the citation data, they are not in the same position as thought
                                                    objRegExpFindEntry.Pattern = objRegExpFindEntry.Pattern & Replace(Mid(objMatchCitationData.value, InStr(objMatchCitationData.value, Chr(34) & " : " & Chr(34)) + 5), Chr(34), "") & ".*"
                                                    If Not blnAuthorsFound Then
                                                        'includes the part to check for "et al."
                                                        objRegExpFindEntry.Pattern = objRegExpFindEntry.Pattern & "((et al\..*)|("
                                                    End If
                                                    'authors were found, we can start searching for the year of publication
                                                    blnAuthorsFound = True
                                                End If
                                            End If
                                        End If
                                    Else
                                        'gets the year of the publication, it is after the entry ends in ("id" : "ITEM-X")
                                        If blnAuthorsFound And Left(objMatchCitationData.value, 5) = "[ [ " & Chr(34) And Right(objMatchCitationData.value, 5) = Chr(34) & " ] ]" Then
                                            strLastYear = Mid(objMatchCitationData.value, 6, Len(objMatchCitationData.value) - 10)
                                            'finishes the pattern including the year and checking if there are more than one author
                                            'if only one author, then removes "et al." from the pattern
                                            If Right(objRegExpFindEntry.Pattern, 2) = "|(" Then
                                                objRegExpFindEntry.Pattern = Left(objRegExpFindEntry.Pattern, Len(objRegExpFindEntry.Pattern) - 14)
                                                objRegExpFindEntry.Pattern = objRegExpFindEntry.Pattern & strLastYear
                                            Else
                                                objRegExpFindEntry.Pattern = objRegExpFindEntry.Pattern & "))" & strLastYear
                                            End If
                                            blnAuthorsFound = False
                                        End If
                                    End If
                                End If

                            Next 'gets the data from current citation entry to build the pattern to find the reference entry in biblio

                            'gets the matches, if any, to check if this reference entry corresponds to the citation being treated
                            If Len(Trim(objMatchCitation.value)) > 6 Then 'includes ", " before the year
                                Set colMatchesFindEntry = objRegExpFindEntry.Execute(objMatchCitation.value)
                            Else
                                Set colMatchesFindEntry = objRegExpFindEntry.Execute(strLastAuthors & ", " & objMatchCitation.value)
                            End If
                            'if this is the corresponding reference entry
                            If colMatchesFindEntry.Count > 0 Then
                                'MsgBox ("Match between DOCUMENT and DATA found:" & vbCrLf & vbCrLf & colMatchesFindEntry.Item(0).Value)
                                Exit For
                            End If

                        Next 'iterates to find all ("id" : "ITEM-X") in colMatchesCitationData to identify where the citation is located


                        'adds the year of current citation entry
                        'we include the year from objMatchCitation (the visible text in the document) because
                        'it may also include a letter in the end (e.g. "2017a") and we need that letter
                        If Mid(objMatchCitation.value, Len(objMatchCitation.value) - 4, 1) = " " Then
                            objRegExpBiblioEntry.Pattern = objRegExpBiblioEntry.Pattern & "\(" & Right(objMatchCitation.value, 4) & "\)"
                        Else
                            objRegExpBiblioEntry.Pattern = objRegExpBiblioEntry.Pattern & "\(" & Right(objMatchCitation.value, 5) & "\)"
                        End If

                        'last verification to make sure we found the citation and not because the for loop reached the end
                        If colMatchesFindEntry.Count = 0 Then
                            'cleans the regular expression as no entries were found
                            objRegExpBiblioEntry.Pattern = "Error: Citation not found"
                        End If

                        'at this point, the regular expression to find the entry in the biblio is ready

                        'initializes the position
                        i = 1
                        'finds the position of the citation entry in the list of references in the biblio
                        blnReferenceEntryFound = False
                        For Each objMatchBiblio In colMatchesBiblio
                            'MsgBox ("Searching for citation in bibliography:" & vbCrLf & vbCrLf & "Using..." & vbCrLf & objRegExpBiblioEntry.Pattern & vbCrLf & objMatchBiblio.Value)
                            'gets the matches, if any, to check if this reference entry corresponds to the citation being treated
                            Set colMatchesBiblioEntry = objRegExpBiblioEntry.Execute(objMatchBiblio.value)
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
                            'MsgBox ("Citation was found in the bibliography" & vbCrLf & vbCrLf & colMatchesBiblioEntry.Item(0).Value)
                            'selects the current field (Mendeley's citation field)
                            sectionField.Select

                            'finds the opening parenthesis (first character of the field),
                            'used to select something inside the field
                            With Selection.Find
                                .Forward = True
                                .Wrap = wdFindStop
                                .Text = "("
                                .Execute
                                blnCitationEntryFound = .found
                            End With

                            'if a match was found (it should always find it, but good practice)
                            'selects the correct entry text from the citation field
                            If blnCitationEntryFound Then
                                'recalculates the selection to include only the match (entry) in citation
                                Selection.MoveEnd Unit:=wdCharacter, Count:=objMatchCitation.FirstIndex + objMatchCitation.Length - 1
                                'if the first character is a blank space
                                If Left(objMatchCitation.value, 1) = " " Then
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
                                Selection.Hyperlinks.Add Anchor:=Selection.range, _
                                    Address:="", SubAddress:="SignetBibliographie_" & format(CStr(i), "00#"), _
                                    ScreenTip:=""

                            End If
                        Else
                            MsgBox "Orphan citation entry found:" & vbCrLf & vbCrLf & _
                                objMatchCitation.value & vbCrLf & vbCrLf & _
                                "Remove it from document!", _
                                vbExclamation, "GAUG_createHyperlinksForCitationsAPA()"
                            'MsgBox ("Orphan citation entry found:" & vbCrLf & vbCrLf & objMatchCitation.Value & vbCrLf & vbCrLf & "Remove it from document!" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Regular expression to find reference in bibliography (from DATA and year from DOCUMENT):" & vbCrLf & vbCrLf & objRegExpBiblioEntry.Pattern & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Last authors (from DOCUMENT):" & vbCrLf & vbCrLf & strLastAuthors & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Year of publication (from DATA):" & vbCrLf & vbCrLf & strLastYear & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Pattern to find matching between DOCUMENT and DATA (DATA):" & vbCrLf & vbCrLf & objRegExpFindEntry.Pattern)
                        End If

                        'Ends: Needs re-work

                        'at this point current citation entry is linked to corresponding reference in biblio

                    Next 'treats all matches (all entries in citation) to generate hyperlinks

                End If 'checks that the string can be compared

            End If 'if it is a citation
        Next 'sectionField

        'at this point all citations are linked to their corresponding reference in biblio

    Next 'documentSection

    'reenables the screen updating
    Application.ScreenUpdating = True

End Sub



'*****************************************************************************************
'*****************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                  **
'**  Last modified: 2019-03-04                                                          **
'**                                                                                     **
'**  Sub GAUG_createHyperlinksForCitationsIEEE()                                        **
'**                                                                                     **
'**  Generates the bookmarks in the bibliography inserted by Mendeley's plugin.         **
'**  Links the citations inserted by Mendeley's plugin to the corresponding entry       **
'**     in the bibliography inserted by Mendeley's plugin.                              **
'**  Generates the hyperlinks for the URLs in the bibliography inserted by              **
'**     Mendeley's plugin.                                                              **
'**  Only for IEEE CSL citation style.                                                  **
'*****************************************************************************************
'*****************************************************************************************
Sub GAUG_createHyperlinksForCitationsIEEE()

    Dim documentSection As Section
    Dim sectionField As Field
    Dim blnFound, blnBibliographyFound, blnReferenceNumberFound, blnCitationNumberFound, blnGenerateHyperlinksForURLs, blnURLFound As Boolean
    Dim intRefereceNumber, intCitationNumber As Integer
    Dim objRegExpCitation, objRegExpURL As RegExp
    Dim colMatchesCitation, colMatchesURL As MatchCollection
    Dim objMatchCitation, objMatchURL As match
    Dim blnIncludeSquareBracketsInHyperlinks As Boolean
    Dim strTypeOfExecution As String
    Dim blnMabEntwickeltSich As Boolean
    Dim stlStyleInDocument As Word.Style
    Dim strStyleForTitleOfBibliography As String
    Dim blnStyleForTitleOfBibliographyFound As Boolean
    Dim strURL As String
    Dim arrNonDetectedURLs, varNonDetectedURL As Variant


'*****************************
'*   Custom configuration    *
'*****************************
    'SEE DOCUMENTATION
    'specifies the name of the font style used for the title of the bibliography
    strStyleForTitleOfBibliography = "Titre de derni�re section"

    'possible values are True and False
    'SEE DOCUMENTATION
    'set to True if the bibliography is in a section with title in style indicated by strStyleForTitleOfBibliography
    blnMabEntwickeltSich = False

    'possible values are True and False
    'SEE DOCUMENTATION
    'if set to True, then the URLs in the bibliography will be converted to hyperlinks
    blnGenerateHyperlinksForURLs = True

    'SEE DOCUMENTATION
    'specifies the URLs, not detected by the regular expression, to generate the hyperlinks in the bibliography
    'note that the last URL does not have a comma after the double quotation mark
    arrNonDetectedURLs = _
        Array( _
            "http://my.first.sample/url/", _
            "https://my.second.sample/url/", _
            "ftp://my.third.sample/url/with/file.pdf" _
            )

    'possible values are True and False
    'SEE DOCUMENTATION
    'if set to True, then argument "RemoveHyperlinks" cannot be used when cleaning old hyperlinks
    blnIncludeSquareBracketsInHyperlinks = False

    'possible values are "RemoveHyperlinks", "CleanEnvironment" and "CleanFullEnvironment"
    'SEE DOCUMENTATION
    strTypeOfExecution = "RemoveHyperlinks"






'*****************************
'*     Initial tasks and     *
'*       verifications       *
'*****************************
    'checks that the style exists in the collection of styles of the document
    For Each stlStyleInDocument In ActiveDocument.Styles
        'checks if the current style is the style for the title of the bibliography
        If stlStyleInDocument = strStyleForTitleOfBibliography Then
            blnStyleForTitleOfBibliographyFound = True
        End If
    Next 'all styles in document

    'if the style was not found
    If blnMabEntwickeltSich And Not blnStyleForTitleOfBibliographyFound Then
        MsgBox "The style " & Chr(34) & strStyleForTitleOfBibliography & Chr(34) & " could not be found." & vbCrLf & vbCrLf & _
            "Add the custom style to Microsoft Word or" & vbCrLf & _
            "set the flag blnMabEntwickeltSich to False." & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_createHyperlinksForCitationsIEEE()"

        'stops the execution
        End
    End If


    'checks for conflicts (square brackets and removal of hyperlinks)
    If blnIncludeSquareBracketsInHyperlinks And strTypeOfExecution = "RemoveHyperlinks" Then
        MsgBox "The hyperlinks will include the square brackets and" & vbCrLf & _
            "Microsoft Word does not like them this way." & vbCrLf & vbCrLf & _
            "You cannot call the macro GAUG_removeHyperlinksForCitations()" & vbCrLf & _
            "with argument " & Chr(34) & "RemoveHyperlinks" & Chr(34) & "." & vbCrLf & _
            "Use " & Chr(34) & "CleanEnvironment" & Chr(34) & " or " & _
            Chr(34) & "CleanFullEnvironment" & Chr(34) & " instead." & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_createHyperlinksForCitationsIEEE()"

        'stops the execution
        End
    End If






'*****************************
'*  Cleaning old hyperlinks  *
'*****************************
    'removes all hyperlinks
    Call GAUG_removeHyperlinksForCitations(strTypeOfExecution)






    'disables the screen updating while creating the hyperlinks to increase speed
    Application.ScreenUpdating = False

'*****************************
'*   Creation of bookmarks   *
'*****************************
    'creates the object for regular expressions (to get all URLs in biblio
    Set objRegExpURL = New RegExp
    'sets the pattern to match every URL in bibliography
    objRegExpURL.Pattern = "((https?)|(ftp)):\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z0-9]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=\[\]\(\)<>;]*)"
    'sets case insensitivity
    objRegExpURL.IgnoreCase = False
    'sets global applicability
    objRegExpURL.Global = True

    'initializes the flag
    blnBibliographyFound = False
    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        'if this is MabEntwickeltSich's document structure
        If blnMabEntwickeltSich Then
            'checks if the section has text with style indicated by strStyleForTitleOfBibliography
            '(it is a section not belonging to any chapter after the sections of parts and chapters)
            blnFound = False
            With documentSection.range.Find
                .Style = strStyleForTitleOfBibliography
                .Execute
                blnFound = .found
            End With
        'this is a document by another user
        Else
            'forces the macro to search for the bibliography in every section
            blnFound = True
        End If

        'checks if the bibliography is in this section
        If blnFound Then
            'checks all fields
            For Each sectionField In documentSection.range.Fields
                'if it is the bibliography
                If sectionField.Type = wdFieldAddin And Trim(sectionField.Code) = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY" Then
                    blnBibliographyFound = True
                    'start the numbering
                    intRefereceNumber = 1
                    Do
                        'selects the current field (Mendeley's bibliography field)
                        sectionField.Select

                        'finds and selects the text of the number of the reference
                        With Selection.Find
                            .Forward = True
                            .Wrap = wdFindStop
                            .Text = "[" & CStr(intRefereceNumber) & "]"
                            .Execute
                            blnReferenceNumberFound = .found
                        End With

                        'if a number of a reference was found, creates the bookmark with the selected text
                        If blnReferenceNumberFound Then
                            'if the square brackets are not part of the hyperlinks
                            If Not blnIncludeSquareBracketsInHyperlinks Then
                                'restricts the selection to only the number
                                With Selection.Find
                                    .Forward = True
                                    .Wrap = wdFindStop
                                    .Text = CStr(intRefereceNumber)
                                    .Execute
                                    blnReferenceNumberFound = .found
                                End With
                            End If

                            'creates the bookmark
                            Selection.Bookmarks.Add _
                                Name:="SignetBibliographie_" & format(CStr(intRefereceNumber), "00#"), _
                                range:=Selection.range
                        End If

                        'continues with the next number
                        intRefereceNumber = intRefereceNumber + 1

                    'while numbers of refereces are found
                    Loop While (blnReferenceNumberFound)

                    'generates the hyperlinks for the URLs in the bibliography if required
                    If blnGenerateHyperlinksForURLs Then

                        'generates the hyperlnks from the list of non detected URLs
                        'the non detected URLs shall be done first or some conflicts may arise
                        For Each varNonDetectedURL In arrNonDetectedURLs
                                'selects the current field (Mendeley's bibliography field)
                                sectionField.Select

                                'finds all instances of current URL
                                Do
                                    'finds and selects the text of the URL
                                    With Selection.Find
                                        .Forward = True
                                        .Wrap = wdFindStop
                                        .Text = CStr(varNonDetectedURL)
                                        .Execute
                                        blnURLFound = .found
                                    End With

                                    'creates the hyperlink
                                    If blnURLFound Then
                                        'checks there is no hyperlink already
                                        If Selection.Hyperlinks.Count = 0 Then
                                            Selection.Hyperlinks.Add Anchor:=Selection.range, _
                                                Address:=Replace(Trim(CStr(varNonDetectedURL)), " ", "%20"), SubAddress:="", _
                                                ScreenTip:=""
                                        End If
                                    End If

                                Loop Until (Not blnURLFound) 'finds all instances of current URL
                        Next 'generates the hyperlnks from the list of non detected URLs

                        'selects the current field (Mendeley's bibliography field)
                        sectionField.Select

                        'checks that the string can be compared (both, Selection and Field.Code)
                        If objRegExpURL.Test(Selection) = True Then
                            'gets the matches (all URLs in the biblio according to the regular expression)
                            Set colMatchesURL = objRegExpURL.Execute(Selection)

                            'treats all matches (all URLs in biblio) to generate hyperlinks
                            For Each objMatchURL In colMatchesURL

                                'removes the last character if a period (".")
                                If Right(objMatchURL.value, 1) = "." Then
                                    strURL = Left(objMatchURL.value, Len(objMatchURL.value) - 1)
                                Else
                                    strURL = objMatchURL.value
                                End If

                                'selects the current field (Mendeley's bibliography field)
                                sectionField.Select

                                'finds all instances of current URL
                                Do
                                    'finds and selects the text of the URL
                                    With Selection.Find
                                        .Forward = True
                                        .Wrap = wdFindStop
                                        .Text = strURL
                                        .Execute
                                        blnURLFound = .found
                                    End With

                                    'creates the hyperlink
                                    If blnURLFound Then
                                        'checks there is no hyperlink already
                                        'this could happen in some cases:
                                        '     when duplicated URLs in bibliography
                                        '     when the hyperlink was created with the list of non detected URLs
                                        '     when a partial URL is detected and the hyperlink was created with the list of non detected URLs
                                        If Selection.Hyperlinks.Count = 0 Then
                                            Selection.Hyperlinks.Add Anchor:=Selection.range, _
                                                Address:=strURL, SubAddress:="", _
                                                ScreenTip:=""
                                        End If
                                    End If

                                Loop Until (Not blnURLFound) 'finds all instances of current URL

                            Next 'treats all matches (all URLs in biblio) to generate hyperlinks
                        End If 'checks that the string can be compared

                    End If 'hyperlinks for URLs in bibliography

                End If 'if it is the biblio
            Next 'sectionField
        End If
    Next 'documentSection


    'if the bibliography could not be located in the document
    If Not blnBibliographyFound Then
        MsgBox "The bibliography could not be located in the document." & vbCrLf & vbCrLf & _
            "Make sure that you have inserted the bibliography via the Mendeley's plugin" & vbCrLf & _
            "and that the custom configuration of the GAUG_* macros is correct." & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_createHyperlinksForCitationsIEEE()"

        'stops the execution
        End
    End If






'*****************************
'*   Linking the bookmarks   *
'*****************************
    'creates the object for regular expressions (to get all entries in current citation, all entries of data in current citation, position of citation entry in biblio)
    Set objRegExpCitation = New RegExp
    'sets the pattern to match every citation entry in current field
    'it should be "[" + Number + "]"
    objRegExpCitation.Pattern = "\[[0-9]+\]"
    'sets case insensitivity
    objRegExpCitation.IgnoreCase = False
    'sets global applicability
    objRegExpCitation.Global = True

    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        'checks all fields
        For Each sectionField In documentSection.range.Fields
            'if it is a citation
            If sectionField.Type = wdFieldAddin And Left(sectionField.Code, 18) = "ADDIN CSL_CITATION" Then

                'selects the current field (Mendeley's citation field)
                sectionField.Select

                'checks that the string can be compared (both, Selection and Field.Code)
                If objRegExpCitation.Test(Selection) = True Then
                    'gets the matches (all entries in the citation according to the regular expression)
                    Set colMatchesCitation = objRegExpCitation.Execute(Selection)

                    'treats all matches (all entries in citation) to generate hyperlinks
                    For Each objMatchCitation In colMatchesCitation
                        'gets the citation number as integer
                        'this will also eliminate leading zeros in numbers (in case of manual modifications)
                        intCitationNumber = CInt(Mid(objMatchCitation.value, 2, Len(objMatchCitation.value) - 2))

                        'to make sure the citation number as text is the same as numeric
                        'and that the citation number is in the bibliography
                        If (("[" & CStr(intCitationNumber) & "]") = objMatchCitation.value) And (intCitationNumber > 0 And intCitationNumber < intRefereceNumber) Then
                            blnCitationNumberFound = True
                        Else
                            blnCitationNumberFound = False
                        End If

                        'if a number of a citation was found (shall always find it), inserts the hyperlink
                        If blnCitationNumberFound Then
                            'selects the current field (Mendeley's citation field)
                            sectionField.Select

                            'finds and selects the text of the number of the reference
                            With Selection.Find
                                .Forward = True
                                .Wrap = wdFindStop
                                .Text = "[" & CStr(intCitationNumber) & "]"
                                .Execute
                                blnReferenceNumberFound = .found
                            End With

                            'if the square brackets are not part of the hyperlinks
                            If Not blnIncludeSquareBracketsInHyperlinks Then
                                'restricts the selection to only the number
                                Selection.MoveStart Unit:=wdCharacter, Count:=1
                                Selection.MoveEnd Unit:=wdCharacter, Count:=-1
                            End If

                            'a cross-reference is not a good idea, it changes the text in citation (or may delete citation):
                            'Selection.Fields.Add Range:=Selection.Range, _
                            '    Type:=wdFieldEmpty, _
                            '    Text:="REF " & Chr(34) & "SignetBibliographie_" & Format(CStr(intCitationNumber), "00#") & Chr(34) & " \h", _
                            '    PreserveFormatting:=True
                            'better to use normal hyperlink:
                            Selection.Hyperlinks.Add Anchor:=Selection.range, _
                                Address:="", SubAddress:="SignetBibliographie_" & format(CStr(intCitationNumber), "00#"), _
                                ScreenTip:=""
                        Else
                            MsgBox "Orphan citation entry found:" & vbCrLf & vbCrLf & _
                                objMatchCitation.value & vbCrLf & vbCrLf & _
                                "Remove it from document!", _
                                vbExclamation, "GAUG_createHyperlinksForCitationsIEEE()"
                        End If
                    Next 'treats all matches (all entries in citation) to generate hyperlinks
                End If 'checks that the string can be compared

            End If 'if it is a citation
        Next 'sectionField
    Next 'documentSection

    'reenables the screen updating
    Application.ScreenUpdating = True

End Sub



'*****************************************************************************************
'*****************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                  **
'**  Last modified: 2019-03-04                                                          **
'**                                                                                     **
'**  Sub GAUG_removeHyperlinksForCitations(strTypeOfExecution As String)                **
'**                                                                                     **
'**  This is an improved version which runs much faster,                                **
'**  but still considered as experimental.                                              **
'**  Make sure you have a backup of your document before you execute it!                **
'**                                                                                     **
'**  Removes the bookmarks generated by GAUG_createHyperlinksForCitations*              **
'**     in the bibliography inserted by Mendeley's plugin.                              **
'**  Removes the hyperlinks generated by GAUG_createHyperlinksForCitations*             **
'**     of the citations inserted by Mendeley's plugin.                                 **
'**  Removes the hyperlinks generated by GAUG_createHyperlinksForCitations*             **
'**     in the bibliography inserted by Mendeley's plugin.                              **
'**  Removes all manual modifications to the citations and bibliography if specified    **
'**                                                                                     **
'**  Parameter strTypeOfExecution can have three different values:                      **
'**  "RemoveHyperlinks":                                                                **
'**     UNEXPECTED RESULTS IF MANUAL MODIFICATIONS EXIST, BUT THE FASTEST               **
'**        Removes the bookmarks and hyperlinks                                         **
'**        Manual modifications to citations and bibliography will remain intact        **
'**  "CleanEnvironment":                                                                **
'**     EXPERIMENTAL, BUT VERY FAST                                                     **
'**        Removes the bookmarks and hyperlinks                                         **
'**        Manual modifications to citations and bibliography will also be removed      **
'**           to have a clean environment                                               **
'**  "CleanFullEnvironment":                                                            **
'**     SAFE, BUT VERY SLOW IN LONG DOCUMENTS                                           **
'**        Removes the bookmarks and hyperlinks                                         **
'**        Manual modifications to citations and bibliography will also be removed      **
'**           to have a clean environment                                               **
'*****************************************************************************************
'*****************************************************************************************
Sub GAUG_removeHyperlinksForCitations(Optional ByVal strTypeOfExecution As String = "RemoveHyperlinks")
    Dim documentSection As Section
    Dim sectionField As Field
    Dim fieldBookmark As Bookmark
    Dim selectionHyperlinks As Hyperlinks
    Dim i As Integer
    Dim blnFound, blnBibliographyFound As Boolean
    Dim sectionFieldName, sectionFieldNewName As String
    Dim objMendeleyApiClient As Object
    Dim cbbUndoEditButton As CommandBarButton
    Dim blnMabEntwickeltSich As Boolean
    Dim stlStyleInDocument As Word.Style
    Dim strStyleForTitleOfBibliography As String
    Dim blnStyleForTitleOfBibliographyFound As Boolean


'*****************************
'*   Custom configuration    *
'*****************************
    'SEE DOCUMENTATION
    'specifies the name of the font style used for the title of the bibliography
    strStyleForTitleOfBibliography = "Titre de derni�re section"

    'possible values are True and False
    'SEE DOCUMENTATION
    'set to True if the bibliography is in a section with title in style indicated by strStyleForTitleOfBibliography
    blnMabEntwickeltSich = False






    'disables the screen updating while removing the hyperlinks to increase speed
    Application.ScreenUpdating = False

'*****************************
'*     Initial tasks and     *
'*       verifications       *
'*****************************
    'checks that the style exists in the collection of styles of the document
    For Each stlStyleInDocument In ActiveDocument.Styles
        'checks if the current style is the style for the title of the bibliography
        If stlStyleInDocument = strStyleForTitleOfBibliography Then
            blnStyleForTitleOfBibliographyFound = True
        End If
    Next 'all styles in document

    'if the style was not found
    If blnMabEntwickeltSich And Not blnStyleForTitleOfBibliographyFound Then
        MsgBox "The style " & Chr(34) & strStyleForTitleOfBibliography & Chr(34) & " could not be found." & vbCrLf & vbCrLf & _
            "Add the custom style to Microsoft Word or" & vbCrLf & _
            "set the flag blnMabEntwickeltSich to False." & vbCrLf & vbCrLf & _
            "Cannot continue removing hyperlinks.", _
            vbCritical, "GAUG_removeHyperlinksForCitations()"

        'stops the execution
        End
    End If


    'selects the type of execution
    Select Case strTypeOfExecution
        Case "RemoveHyperlinks"
            'nothing to do here
        Case "CleanEnvironment"
            'get the API Client from Mendeley
            Set objMendeleyApiClient = Application.Run("Mendeley.mendeleyApiClient") 'MabEntwickeltSich: This is the way to call the macro directly from Mendeley
        Case "CleanFullEnvironment"
            'gets the Undo Edit Button
            Set cbbUndoEditButton = Application.Run("MendeleyLib.getUndoEditButton") 'MabEntwickeltSich: This is the way to call the macro directly from Mendeley
        Case Else
            'reenables the screen updating
            Application.ScreenUpdating = True

            MsgBox "Unknown execution type " & Chr(34) & strTypeOfExecution & Chr(34) & " for GAUG_removeHyperlinksForCitations()." & vbCrLf & vbCrLf & _
                "Please, correct the error and try again.", _
                vbCritical, "GAUG_removeHyperlinksForCitations()"

            'the execution option is not correct
            End
        End Select






'*****************************
'*  Cleaning old hyperlinks  *
'*****************************
    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        For Each sectionField In documentSection.range.Fields
            'if it is a citation
            If sectionField.Type = wdFieldAddin And Left(sectionField.Code, 18) = "ADDIN CSL_CITATION" Then
                sectionField.Select

                'selects the type of execution to remove hyperlinks
                Select Case strTypeOfExecution
                    Case "RemoveHyperlinks"
                        'gets all hyperlinks from selection
                        Set selectionHyperlinks = Selection.Hyperlinks

                        'deletes all hyperlinks
                        For i = selectionHyperlinks.Count To 1 Step -1
                            'this method produces errors if the hyperlinks include the square brackets in IEEE
                            If Left(selectionHyperlinks(1).range.Text, 1) = "[" Then
                                MsgBox "There was an error removing the hyperlinks" & vbCrLf & _
                                    "because they include the square brackets and" & vbCrLf & _
                                    "Microsoft Word does not like them this way." & vbCrLf & vbCrLf & _
                                    "You cannot call the macro GAUG_removeHyperlinksForCitations()" & vbCrLf & _
                                    "with argument " & Chr(34) & "RemoveHyperlinks" & Chr(34) & "." & vbCrLf & _
                                    "Use " & Chr(34) & "CleanEnvironment" & Chr(34) & " or " & _
                                    Chr(34) & "CleanFullEnvironment" & Chr(34) & " instead." & vbCrLf & _
                                    "You can also call the respective wrapper function." & vbCrLf & vbCrLf & _
                                    "Cannot continue removing hyperlinks.", _
                                    vbCritical, "GAUG_removeHyperlinksForCitations()"

                                'reenables the screen updating
                                Application.ScreenUpdating = True

                                'stops the execution
                                End
                            End If

                            'deletes the current hyperlink
                            selectionHyperlinks(1).Delete
                        Next
                    Case "CleanEnvironment"
                        'copied from Mendeley.undoEdit(), but removing the code that updates the toolbar in Microsoft Word (making the original function very slow)
                        'restores the citations to the original state (deletes hyperlinks)
                        sectionFieldName = Application.Run("ZoteroLib.getMarkName", sectionField)
                        sectionFieldNewName = objMendeleyApiClient.undoManualFormat(sectionFieldName)
                        Call Application.Run("ZoteroLib.fnRenameMark", sectionField, sectionFieldNewName) 'MabEntwickeltSich: This is another way to call the macro directly from Mendeley
                        Call Application.Run("ZoteroLib.subSetMarkText", sectionField, INSERT_CITATION_TEXT) 'MabEntwickeltSich: This is another way to call the macro directly from Mendeley
                    Case "CleanFullEnvironment"
                        'restores the citations to the original state (deletes hyperlinks)
                        'slow version
                        cbbUndoEditButton.Execute
                    End Select

            End If
        Next
    Next






'*****************************
'*  Cleaning old bookmarks   *
'*****************************
    'initializes the flag
    blnBibliographyFound = False
    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        'if this is MabEntwickeltSich's document structure
        If blnMabEntwickeltSich Then
            'checks if the section has text with style indicated by strStyleForTitleOfBibliography
            '(it is a section not belonging to any chapter after the sections of parts and chapters)
            blnFound = False
            With documentSection.range.Find
                .Style = strStyleForTitleOfBibliography
                .Execute
                blnFound = .found
            End With
        'this is a document by another user
        Else
            'forces the macro to search for the bibliography in every section
            blnFound = True
        End If

        'checks if the bibliography is in this section
        If blnFound Then
            For Each sectionField In documentSection.range.Fields
                'if it is the bibliography
                If sectionField.Type = wdFieldAddin And Trim(sectionField.Code) = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY" Then
                    blnBibliographyFound = True
                    sectionField.Select
                    'deletes all bookmarks
                    For Each fieldBookmark In Selection.Bookmarks
                        'deletes current bookmark
                        fieldBookmark.Delete
                    Next

                    sectionField.Select
                    'gets all URL hyperlinks from selection
                    Set selectionHyperlinks = Selection.Hyperlinks
                    'deletes all URL hyperlinks
                    'MsgBox "Total number of hyperlinks in biblio: " & CStr(selectionHyperlinks.Count)
                    For i = selectionHyperlinks.Count To 1 Step -1
                        'deletes the current URL hyperlink
                        selectionHyperlinks(1).Delete
                    Next
                End If
            Next
        End If

    Next


    'if the bibliography could not be located in the document
    If Not blnBibliographyFound Then
        MsgBox "The bibliography could not be located in the document." & vbCrLf & vbCrLf & _
            "Make sure that you have inserted the bibliography via the Mendeley's plugin" & vbCrLf & _
            "and that the custom configuration of the GAUG_* macros is correct." & vbCrLf & vbCrLf & _
            "Cannot continue removing hyperlinks.", _
            vbCritical, "GAUG_removeHyperlinksForCitations()"

        'stops the execution
        End
    End If


    'selects the type of execution
    Select Case strTypeOfExecution
        Case "RemoveHyperlinks"
            'nothing to do here
        Case "CleanEnvironment"
            'copied from Mendeley.undoEdit(), but removing the code that updates the toolbar in Microsoft Word (making the original function very slow)
            Call Application.Run("MendeleyLib.refreshDocument") 'MabEntwickeltSich: This is another way to call the macro directly from Mendeley
        Case "CleanFullEnvironment"
            'nothing to do here
        End Select

    'reenables the screen updating
    Application.ScreenUpdating = True

End Sub



'*****************************************************************************************
'*****************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                  **
'**  Last modified: 2017-01-11                                                          **
'**                                                                                     **
'**  Sub GAUG_removeHyperlinks()                                                        **
'**                                                                                     **
'**  Calls Sub GAUG_removeHyperlinksForCitations(strTypeOfExecution As String)          **
'**     with parameter strTypeOfExecution = "RemoveHyperlinks"                          **
'*****************************************************************************************
'*****************************************************************************************
Sub GAUG_removeHyperlinks()
    'removes all bookmarks and hyperlinks from the citations and bibliography
    Call GAUG_removeHyperlinksForCitations("RemoveHyperlinks")
End Sub



'*****************************************************************************************
'*****************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                  **
'**  Last modified: 2017-01-11                                                          **
'**                                                                                     **
'**  Sub GAUG_cleanEnvironment()                                                        **
'**                                                                                     **
'**  Calls Sub GAUG_removeHyperlinksForCitations(strTypeOfExecution As String)          **
'**     with parameter strTypeOfExecution = "CleanEnvironment"                          **
'*****************************************************************************************
'*****************************************************************************************
Sub GAUG_cleanEnvironment()
    'removes all bookmarks, hyperlinks and manual modifications to the citations and bibliography
    Call GAUG_removeHyperlinksForCitations("CleanEnvironment")
End Sub



'*****************************************************************************************
'*****************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                  **
'**  Last modified: 2017-01-11                                                          **
'**                                                                                     **
'**  Sub GAUG_cleanFullEnvironment()                                                    **
'**                                                                                     **
'**  Calls Sub GAUG_removeHyperlinksForCitations(strTypeOfExecution As String)          **
'**     with parameter strTypeOfExecution = "CleanFullEnvironment"                      **
'*****************************************************************************************
'*****************************************************************************************
Sub GAUG_cleanFullEnvironment()
    'removes all bookmarks, hyperlinks and manual modifications to the citations and bibliography
    Call GAUG_removeHyperlinksForCitations("CleanFullEnvironment")
End Sub



