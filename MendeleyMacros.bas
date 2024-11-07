Attribute VB_Name = "MendeleyMacros"
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                                                                        **
'**  Last modified: 2024-10-02                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getAvailableMendeleyVersion(Optional intUseMendeleyVersion As Integer = 0) As Integer                                      **
'**                                                                                                                                           **
'**  Finds the available version of Mendeley.                                                                                                 **
'**  The version can be overridden by the optional parameter intUseMendeleyVersion,                                                           **
'**     useful if the macros fail to detect the correct version.                                                                              **
'**                                                                                                                                           **
'**  Parameter intUseMendeleyVersion can have three different values:                                                                         **
'**  0: (DEFAULT)                                                                                                                             **
'**     Autodetect Mendeley's version                                                                                                         **
'**  1:                                                                                                                                       **
'**     Use version 1.x of Mendeley Desktop                                                                                                   **
'**  2:                                                                                                                                       **
'**     Use version 2.x of Mendeley Reference Manager                                                                                         **
'**                                                                                                                                           **
'**  RETURNS: An integer with the major version number of Mendeley that is installed.                                                         **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_getAvailableMendeleyVersion(Optional ByVal intUseMendeleyVersion As Integer = 0) As Integer

    Dim intAvailableMendeleyVersion As Integer
    Dim blnFound As Boolean
    Dim fldField As Field
    Dim ccContentControl As ContentControl


    'if the optional argument is not within valid versions
    If intUseMendeleyVersion < 0 Or intUseMendeleyVersion > 2 Then
        MsgBox "The version " & intUseMendeleyVersion & " of Mendeley's plugin is not valid." & vbCrLf & vbCrLf & _
            "Use version 1 or version 2 instead," & vbCrLf & _
            "or 0 for the macros to automatically detect it." & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getAvailableMendeleyVersion(intUseMendeleyVersion)"

        'stops the execution
        End
    End If


    'if the macros must automatically detect the version of Mendeley's plugin
    If intUseMendeleyVersion = 0 Then
        blnFound = False
        'checks all fields in the document
        For Each fldField In ActiveDocument.Fields
            'if this is a citation by Mendeley Desktop 1.x
            If fldField.Type = wdFieldAddin And Left(fldField.Code, 18) = "ADDIN CSL_CITATION" Then
                blnFound = True
                'sets available version to 1
                intAvailableMendeleyVersion = 1
                'exits loop
                Exit For
            End If
        Next fldField

        'if version 1.x was not found, tries to detect if 2.x is installed
        If Not blnFound Then
            'checks all content controls in the document
            For Each ccContentControl In ActiveDocument.ContentControls
                'if this is a citation by Mendeley Reference Manager 2.x
                If ccContentControl.Type = wdContentControlRichText And Left(Trim(ccContentControl.Tag), 21) = "MENDELEY_CITATION_v3_" Then
                    blnFound = True
                    'sets available version to 2
                    intAvailableMendeleyVersion = 2
                    'exits loop
                    Exit For
                End If
            Next ccContentControl
        End If

        'if the macros could not detect the version of Mendeley's plugin
        '(there are no citations added by Mendeley in the entire document)
        If Not blnFound Then
            MsgBox "The version of Mendeley's plugin could not be detected." & vbCrLf & vbCrLf & _
                "Cannot continue creating hyperlinks.", _
                vbCritical, "GAUG_getAvailableMendeleyVersion(intUseMendeleyVersion)"

            'stops the execution
            End
        End If

    'if the user has specified the version of Mendeley's plugin
    Else
        'overrides the available version
        intAvailableMendeleyVersion = intUseMendeleyVersion
    End If

    'returns the available version of Mendeley
    GAUG_getAvailableMendeleyVersion = intAvailableMendeleyVersion

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                                                                        **
'**  Last modified: 2024-10-02                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getAllCitationsFullInformation(intMendeleyVersion As Integer) As String                                                    **
'**                                                                                                                                           **
'**  Finds and returns the full information of all citations, when available.                                                                 **
'**                                                                                                                                           **
'**  Parameter intMendeleyVersion can have two different values:                                                                              **
'**  1:                                                                                                                                       **
'**     Use version 1.x of Mendeley Desktop                                                                                                   **
'**        The function returns an empty string due to the fact that                                                                          **
'**        Mendeley Desktop stores the information of each citation inside the field of the citation                                          **
'**  2:                                                                                                                                       **
'**     Use version 2.x of Mendeley Reference Manager                                                                                         **
'**        The function returns the information of all citations in a single string                                                           **
'**                                                                                                                                           **
'**  RETURNS: A string that contains all the information of all citations (when Mendeley Reference Manager 2.x is available).                 **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_getAllCitationsFullInformation(ByVal intMendeleyVersion As Integer) As String

    Dim objRegExpWordOpenXMLCitations As RegExp
    Dim colMatchesWordOpenXMLCitations As MatchCollection
    Dim objMatchWordOpenXMLCitations As match
    Dim strAllCitationsFullInformation As String
    Dim blnFound As Boolean


    'if the argument is not within valid versions
    If intMendeleyVersion < 1 Or intMendeleyVersion > 2 Then
        MsgBox "The version " & intMendeleyVersion & " of Mendeley's plugin is not valid." & vbCrLf & vbCrLf & _
            "Use version 1 or version 2 instead," & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getAllCitationsFullInformation(intMendeleyVersion)"

        'stops the execution
        End
    End If


    'initialize the flag
    blnFound = False


    Select Case intMendeleyVersion
        'Mendeley Desktop 1.x is installed
        Case 1
            blnFound = True
            'nothing to do here, Mendeley Desktop 1.x does not provide the information of all citations as a single block of information
            strAllCitationsFullInformation = ""

        'Mendeley Reference Manager 2.x is installed
        Case 2
            Set objRegExpWordOpenXMLCitations = New RegExp
            'ActiveDocument.WordOpenXML contains everything on the document, including hidden information about the citations added by Mendeley's plugin
            'we need that hidden information to be able to match every citation to the corresponding entry in the bibliography
            '(the information is stored as one block of data within WordOpenXML)
            'sets the pattern to match everything from '<we:property name="MENDELEY_CITATIONS"' to (but not including) '<we:property name="MENDELEY_CITATIONS_STYLE"' or '</we:properties><we:bindings/>'
            'for details on the regular expression, see https://stackoverflow.com/questions/7124778/how-can-i-match-anything-up-until-this-sequence-of-characters-in-a-regular-exp
            objRegExpWordOpenXMLCitations.Pattern = "<we:property name=\" & Chr(34) & "MENDELEY_CITATIONS\" & Chr(34) & ".+?(?=((<we:property name=\" & Chr(34) & "MENDELEY_CITATIONS_STYLE\" & Chr(34) & ")|(</we:properties><we:bindings/>)))"
            'sets case insensitivity
            objRegExpWordOpenXMLCitations.IgnoreCase = False
            'sets global applicability
            objRegExpWordOpenXMLCitations.Global = True

            'checks that the string can be compared
            If (objRegExpWordOpenXMLCitations.Test(ActiveDocument.WordOpenXML) = True) Then
                'gets the matches (all information of citations according to the regular expression)
                Set colMatchesWordOpenXMLCitations = objRegExpWordOpenXMLCitations.Execute(ActiveDocument.WordOpenXML)

                'there should be only one match which contains the information of all citations
                If colMatchesWordOpenXMLCitations.Count = 1 Then
                    blnFound = True
                    'treats all matches (the only one)
                    For Each objMatchWordOpenXMLCitations In colMatchesWordOpenXMLCitations
                        'replaces all '&quot;' by '"' to handle the string more easily
                        strAllCitationsFullInformation = Replace(objMatchWordOpenXMLCitations.value, "&quot;", Chr(34))
                    Next objMatchWordOpenXMLCitations
                End If
            End If
    End Select


    'if no information could be found (or many matches where found which is not expected)
    If Not blnFound Then
        MsgBox "Could not find ONLY one match when looking for the full information of all citations." & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getAllCitationsFullInformation(intMendeleyVersion)"

        'stops the execution
        End
    End If


    'returns the string that contains all the information of all citations (when Mendeley Reference Manager 2.x is available)
    GAUG_getAllCitationsFullInformation = strAllCitationsFullInformation

End Function



'*****************************************************************************************************************************************************************************************
'*****************************************************************************************************************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                                                                                                                  **
'**  Last modified: 2024-10-02                                                                                                                                                          **
'**                                                                                                                                                                                     **
'**  GAUG_getCitationFullInfo(intMendeleyVersion As Integer, strAllCitationsFullInformation As String, fldCitation As Field, ccCitation As ContentControl) As String                    **
'**                                                                                                                                                                                     **
'**  Finds and returns the full information of a particular citation.                                                                                                                   **
'**                                                                                                                                                                                     **
'**  Parameter intMendeleyVersion can have two different values:                                                                                                                        **
'**  1:                                                                                                                                                                                 **
'**     Use version 1.x of Mendeley Desktop                                                                                                                                             **
'**  2:                                                                                                                                                                                 **
'**     Use version 2.x of Mendeley Reference Manager                                                                                                                                   **
'**  Parameter strAllCitationsFullInformation is a string that contains all information of all citations when Mendeley Reference Manager 2.x is used                                    **
'**  Parameter fldCitation is the citation's field when Mendeley Desktop 1.x is used                                                                                                    **
'**  Parameter ccCitation is the citation's contet control when Mendeley Reference Manager 2.x is used                                                                                  **
'**                                                                                                                                                                                     **
'**  RETURNS: A string that contains all the information of the citation.                                                                                                               **
'*****************************************************************************************************************************************************************************************
'*****************************************************************************************************************************************************************************************
Function GAUG_getCitationFullInfo(ByVal intMendeleyVersion As Integer, ByVal strAllCitationsFullInformation As String, ByVal fldCitation As Field, ByVal ccCitation As ContentControl) As String

    Dim objRegExpVisibleCitationItems As RegExp
    Dim colMatchesVisibleCitationItems As MatchCollection
    Dim objMatchVisibleCitationItem As match
    Dim strCitationFullInfo As String
    Dim blnFound As Boolean


    'if the argument is not within valid versions
    If intMendeleyVersion < 1 Or intMendeleyVersion > 2 Then
        MsgBox "The version " & intMendeleyVersion & " of Mendeley's plugin is not valid." & vbCrLf & vbCrLf & _
            "Use version 1 or version 2 instead," & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getCitationFullInfo(intMendeleyVersion, strAllCitationsFullInformation, fldCitation, ccCitation)"

        'stops the execution
        End
    End If


    'initialize the flag
    blnFound = False
    'initializes the variable
    strCitationFullInfo = ""

    Select Case intMendeleyVersion
        'Mendeley Desktop 1.x is installed
        Case 1
            'if the citation's field is not empty
            If Not (fldCitation Is Nothing) Then
                blnFound = True
                'the full information of the cition is inside the citation's field
                strCitationFullInfo = fldCitation.Code
            End If

        'Mendeley Reference Manager 2.x is installed
        Case 2
            'if the citation's content control is not empty
            If Not (ccCitation Is Nothing) Then

                Set objRegExpVisibleCitationItems = New RegExp
                'sets the pattern to match everything from '"citationID":"MENDELEY_CITATION_' to (but not including) '"citationID":"MENDELEY_CITATION_' or to the end of the string
                'this gets individual citations, then the correct one can be selected
                objRegExpVisibleCitationItems.Pattern = "{\" & Chr(34) & "citationID\" & Chr(34) & ":\" & Chr(34) & "MENDELEY_CITATION_.+?(?=(({\" & Chr(34) & "citationID\" & Chr(34) & ":\" & Chr(34) & "MENDELEY_CITATION_)|($)))"
                'sets case insensitivity
                objRegExpVisibleCitationItems.IgnoreCase = False
                'sets global applicability
                objRegExpVisibleCitationItems.Global = True

                'checks that the string can be compared
                If (objRegExpVisibleCitationItems.Test(strAllCitationsFullInformation) = True) Then
                    'gets the matches (all information of individual citations according to the regular expression)
                    Set colMatchesVisibleCitationItems = objRegExpVisibleCitationItems.Execute(strAllCitationsFullInformation)

                    'treats all matches (all individual citations) to find the correct one
                    For Each objMatchVisibleCitationItem In colMatchesVisibleCitationItems
                        'if the tag of the searched citation is in the current match
                        If InStr(1, objMatchVisibleCitationItem.value, ccCitation.Tag, 1) > 0 Then
                            blnFound = True
                            'the full information of the cition is in this match
                            strCitationFullInfo = objMatchVisibleCitationItem.value
                            'exits loop
                            Exit For
                        End If
                    Next objMatchVisibleCitationItem
                End If 'checks that the string can be compared

            End If 'if the citation's content control is not empty
    End Select


    'if no information could be found
    If Not blnFound Then
        MsgBox "Could not find the full information of the citation." & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getCitationFullInfo(intMendeleyVersion, strAllCitationsFullInformation, fldCitation, ccCitation)"

        'stops the execution
        End
    End If


    'returns the citation information
    GAUG_getCitationFullInfo = strCitationFullInfo

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                                                                        **
'**  Last modified: 2024-10-02                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getCitationItemsFromCitationFullInfo(ByVal intMendeleyVersion As Integer,                                                  **
'**     ByVal strCitationFullInfo As String) As Variant()                                                                                     **
'**                                                                                                                                           **
'**  Returns the full information of all the individual items of a particular citation.                                                       **
'**                                                                                                                                           **
'**  Parameter intMendeleyVersion can have two different values:                                                                              **
'**  1:                                                                                                                                       **
'**     Use version 1.x of Mendeley Desktop                                                                                                   **
'**  2:                                                                                                                                       **
'**     Use version 2.x of Mendeley Reference Manager                                                                                         **
'**  Parameter strCitationFullInfo is a string that contains the full information of the citation                                             **
'**                                                                                                                                           **
'**  RETURNS: An array that contains the full information of all the individual items of the citation.                                        **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_getCitationItemsFromCitationFullInfo(ByVal intMendeleyVersion As Integer, ByVal strCitationFullInfo As String) As Variant()

    Dim varCitationItemsFromCitationFullInfo() As Variant
    Dim intTotalCitationItems As Integer

    Dim objRegExpCitationItems As RegExp
    Dim colMatchesCitationItems As MatchCollection
    Dim objMatchVisibleCitationItemItem As match


    'if the argument is not within valid versions
    If intMendeleyVersion < 1 Or intMendeleyVersion > 2 Then
        MsgBox "The version " & intMendeleyVersion & " of Mendeley's plugin is not valid." & vbCrLf & vbCrLf & _
            "Use version 1 or version 2 instead," & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getCitationItemsFromCitationFullInfo(intMendeleyVersion, strCitationFullInfo)"

        'stops the execution
        End
    End If


    'if the citation's full info is not empty
    If Not (strCitationFullInfo = "") Then

        Set objRegExpCitationItems = New RegExp
        'sets case insensitivity
        objRegExpCitationItems.IgnoreCase = False
        'sets global applicability
        objRegExpCitationItems.Global = True

        'initializes the counter
        intTotalCitationItems = 0

        'builds the regular expression according to the version of Mendeley
        Select Case intMendeleyVersion
            'Mendeley Desktop 1.x is installed
            Case 1
                    'sets the pattern to match everything from '{"id":"ITEM' to (but not including) ',{"id":"ITEM' or to the end of the string if no other item is found
                    'this gets individual citation items
                    objRegExpCitationItems.Pattern = "{\s*\" & Chr(34) & "id\" & Chr(34) & "\s*:\s*\" & Chr(34) & "ITEM.+?(?=((,\s*{\s*\" & Chr(34) & "id\" & Chr(34) & "\s*:\s*\" & Chr(34) & "ITEM)|($)))"

            'Mendeley Reference Manager 2.x is installed
            Case 2
                    'sets the pattern to match everything from '{"id":"' to (but not including) ',{"id":"' or to the end of the string if no other item is found
                    'this gets individual citation items
                    objRegExpCitationItems.Pattern = "{\s*\" & Chr(34) & "id\" & Chr(34) & "\s*:\s*\" & Chr(34) & ".+?(?=((,\s*{\s*\" & Chr(34) & "id\" & Chr(34) & "\s*:\s*\" & Chr(34) & ")|($)))"
        End Select


        'checks that the string can be compared
        If (objRegExpCitationItems.Test(strCitationFullInfo) = True) Then
            'gets the matches (individual citation items according to the regular expression)
            Set colMatchesCitationItems = objRegExpCitationItems.Execute(strCitationFullInfo)

            'treats all matches (all individual citation items)
            For Each objMatchVisibleCitationItemItem In colMatchesCitationItems
                'MsgBox objMatchVisibleCitationItemItem.value
                'updates the counter to include this citation item
                intTotalCitationItems = intTotalCitationItems + 1
                'adds the full information of the citation item to the list
                ReDim Preserve varCitationItemsFromCitationFullInfo(1 To intTotalCitationItems)
                varCitationItemsFromCitationFullInfo(intTotalCitationItems) = objMatchVisibleCitationItemItem.value
            Next objMatchVisibleCitationItemItem
        End If 'checks that the string can be compared

    End If 'if the citation's full info is not empty


    'returns the list of all items (individual citations within the field or content control) from the citation full information
    GAUG_getCitationItemsFromCitationFullInfo = varCitationItemsFromCitationFullInfo

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                                                                        **
'**  Last modified: 2024-10-02                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getAuthorsEditorsFromCitationItem(ByVal intMendeleyVersion As Integer, ByVal strAuthorEditor As String,                    **
'**     ByVal strCitationItem As String) As Variant()                                                                                         **
'**                                                                                                                                           **
'**  Returns the list of authors or editors of the individual citation item.                                                                  **
'**                                                                                                                                           **
'**  Parameter intMendeleyVersion can have two different values:                                                                              **
'**  1:                                                                                                                                       **
'**     Use version 1.x of Mendeley Desktop                                                                                                   **
'**  2:                                                                                                                                       **
'**     Use version 2.x of Mendeley Reference Manager                                                                                         **
'**  Parameter strAuthorEditor can have two different values:                                                                                 **
'**  "author":                                                                                                                                **
'**     Retrieve the authors of the individual citation item                                                                                  **
'**  "editor":                                                                                                                                **
'**     Retrieve the editors of the individual citation item                                                                                  **
'**  Parameter strCitationItem is a string that contains the full information of the citation item                                            **
'**                                                                                                                                           **
'**  RETURNS: An array that contains the list of authors or editors of the citation item.                                                     **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_getAuthorsEditorsFromCitationItem(ByVal intMendeleyVersion As Integer, ByVal strAutorEditor As String, ByVal strCitationItem As String) As Variant()

    Dim varAuthorsFomCitationItem() As Variant
    Dim intTotalAuthorsEditorsFromCitationItem As Integer
    Dim strFamilyName As String

    Dim objRegExpAuthorsFromCitationItem, objRegExpAuthorFamilyNamesFromCitationItem As RegExp
    Dim colMatchesAuthorsFromCitationItem, colMatchesAuthorFamilyNamesFromCitationItem As MatchCollection
    Dim objMatchAuthorFromCitationItem, objMatchAuthorFamilyNameFromCitationItem As match


    'if the argument is not within valid versions
    If intMendeleyVersion < 1 Or intMendeleyVersion > 2 Then
        MsgBox "The version " & intMendeleyVersion & " of Mendeley's plugin is not valid." & vbCrLf & vbCrLf & _
            "Use version 1 or version 2 instead," & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getAuthorsEditorsFromCitationItem(intMendeleyVersion, strCitationItem)"

        'stops the execution
        End
    End If


    'if the citation item's full info is not empty
    If Not (strCitationItem = "") Then

        Set objRegExpAuthorsFromCitationItem = New RegExp
        'sets case insensitivity
        objRegExpAuthorsFromCitationItem.IgnoreCase = False
        'sets global applicability
        objRegExpAuthorsFromCitationItem.Global = True

        'builds the regular expression according to the version of Mendeley
        Select Case intMendeleyVersion
            'Mendeley Desktop 1.x is installed
            Case 1
                    'sets the pattern to match everything from '"author":[' or '"editor":[' to (but not including) ']'
                    'this gets the full list of authors from the citation item
                    objRegExpAuthorsFromCitationItem.Pattern = "\" & Chr(34) & strAutorEditor & "\" & Chr(34) & "\s*:\s*\[.+?(?=\])"

            'Mendeley Reference Manager 2.x is installed
            Case 2
                    'sets the pattern to match everything from '"author":[' or '"editor":[' to (but not including) ']'
                    'this gets the full list of authors from the citation item
                    objRegExpAuthorsFromCitationItem.Pattern = "\" & Chr(34) & strAutorEditor & "\" & Chr(34) & "\s*:\s*\[.+?(?=\])"
        End Select


        'checks that the string can be compared
        If (objRegExpAuthorsFromCitationItem.Test(strCitationItem) = True) Then
            'gets the matches (list of all authors as a single block of data)
            Set colMatchesAuthorsFromCitationItem = objRegExpAuthorsFromCitationItem.Execute(strCitationItem)

            'treats all matches (there should be at most one match, zero when editors are listed instead of the authors)
            For Each objMatchAuthorFromCitationItem In colMatchesAuthorsFromCitationItem

                Set objRegExpAuthorFamilyNamesFromCitationItem = New RegExp
                'sets case insensitivity
                objRegExpAuthorFamilyNamesFromCitationItem.IgnoreCase = False
                'sets global applicability
                objRegExpAuthorFamilyNamesFromCitationItem.Global = True

                'initializes the counter
                intTotalAuthorsEditorsFromCitationItem = 0

                'builds the regular expression according to the version of Mendeley
                Select Case intMendeleyVersion
                    'Mendeley Desktop 1.x is installed
                    Case 1
                            'sets the pattern to match everything from '"family":"' to (but not including) '"'
                            'this gets the family names of authors from the citation item
                            objRegExpAuthorFamilyNamesFromCitationItem.Pattern = "\" & Chr(34) & "family\" & Chr(34) & "\s*:\s*\" & Chr(34) & ".+?(?=\" & Chr(34) & ")"

                    'Mendeley Reference Manager 2.x is installed
                    Case 2
                            'sets the pattern to match everything from '{"family":"' to (but not including) '"'
                            'this gets the family names of authors from the citation item
                            objRegExpAuthorFamilyNamesFromCitationItem.Pattern = "{\s*\" & Chr(34) & "family\" & Chr(34) & "\s*:\s*\" & Chr(34) & ".+?(?=\" & Chr(34) & ")"
                End Select


                'checks that the string can be compared
                If (objRegExpAuthorFamilyNamesFromCitationItem.Test(objMatchAuthorFromCitationItem.value) = True) Then
                    'gets the matches (the family names of all authors)
                    Set colMatchesAuthorFamilyNamesFromCitationItem = objRegExpAuthorFamilyNamesFromCitationItem.Execute(objMatchAuthorFromCitationItem.value)

                    'treats all matches (the family name of the authors, if any)
                    For Each objMatchAuthorFamilyNameFromCitationItem In colMatchesAuthorFamilyNamesFromCitationItem
                        'gets only the family name, without the extra characters in the match
                        'from '{"family":"FamilyName' to just "FamilyName"
                        strFamilyName = Right(objMatchAuthorFamilyNameFromCitationItem.value, Len(objMatchAuthorFamilyNameFromCitationItem.value) - InStr(objMatchAuthorFamilyNameFromCitationItem.value, ":") - 1)
                        'updates the counter to include this family name of the author
                        intTotalAuthorsEditorsFromCitationItem = intTotalAuthorsEditorsFromCitationItem + 1
                        'adds the family name of the author to the list
                        ReDim Preserve varAuthorsFomCitationItem(1 To intTotalAuthorsEditorsFromCitationItem)
                        varAuthorsFomCitationItem(intTotalAuthorsEditorsFromCitationItem) = strFamilyName
                    Next objMatchAuthorFamilyNameFromCitationItem
                End If 'checks that the string can be compared


            Next objMatchAuthorFromCitationItem 'treats all matches (there should be at most one match, zero when editors are listed instead of the authors)

        End If 'checks that the string can be compared

    End If 'if the citation's full info is not empty


    'returns the list of the family names of the authors
    GAUG_getAuthorsEditorsFromCitationItem = varAuthorsFomCitationItem

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                                                                        **
'**  Last modified: 2024-10-02                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getYearFromCitationItem(ByVal intMendeleyVersion As Integer, ByVal strCitationItem As String) As String                    **
'**                                                                                                                                           **
'**  Returns the year of issue of the individual citation item.                                                                               **
'**                                                                                                                                           **
'**  Parameter intMendeleyVersion can have two different values:                                                                              **
'**  1:                                                                                                                                       **
'**     Use version 1.x of Mendeley Desktop                                                                                                   **
'**  2:                                                                                                                                       **
'**     Use version 2.x of Mendeley Reference Manager                                                                                         **
'**  Parameter strCitationItem is a string that contains the full information of the citation item                                            **
'**                                                                                                                                           **
'**  RETURNS: A string with the year of issue of the citation item (it does not include the letter that may be present after the year).       **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_getYearFromCitationItem(ByVal intMendeleyVersion As Integer, ByVal strCitationItem As String) As String

    Dim strYearFromCitationItem As String

    Dim objRegExpYearFromCitationItem As RegExp
    Dim colMatchesYearFromCitationItem As MatchCollection
    Dim objMatchYearFromCitationItem As match


    'if the argument is not within valid versions
    If intMendeleyVersion < 1 Or intMendeleyVersion > 2 Then
        MsgBox "The version " & intMendeleyVersion & " of Mendeley's plugin is not valid." & vbCrLf & vbCrLf & _
            "Use version 1 or version 2 instead," & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getYearFromCitationItem(intMendeleyVersion, strCitationItem)"

        'stops the execution
        End
    End If


    'if the citation item's full info is not empty
    If Not (strCitationItem = "") Then

        Set objRegExpYearFromCitationItem = New RegExp
        'sets case insensitivity
        objRegExpYearFromCitationItem.IgnoreCase = False
        'sets global applicability
        objRegExpYearFromCitationItem.Global = True

        'builds the regular expression according to the version of Mendeley
        Select Case intMendeleyVersion
            'Mendeley Desktop 1.x is installed
            Case 1
                    'sets the pattern to match everything from '"issued":{"date-parts":[[' to (but not including) ']]' or ','
                    'this gets only the year from the citation item, skips the month if present
                    objRegExpYearFromCitationItem.Pattern = "\" & Chr(34) & "issued\" & Chr(34) & "\s*:\s*{\s*\" & Chr(34) & "date\-parts\" & Chr(34) & "\s*:\s*\[\s*\[.+?(?=((\s*\]\s*\])|(,)))"

            'Mendeley Reference Manager 2.x is installed
            Case 2
                    'sets the pattern to match everything from '"issued":{"date-parts":[[' to (but not including) ']]' or ','
                    'this gets only the year from the citation item, skips the month if present
                    objRegExpYearFromCitationItem.Pattern = "\" & Chr(34) & "issued\" & Chr(34) & "\s*:\s*{\s*\" & Chr(34) & "date\-parts\" & Chr(34) & "\s*:\s*\[\s*\[.+?(?=((\s*\]\s*\])|(,)))"
        End Select


        'checks that the string can be compared
        If (objRegExpYearFromCitationItem.Test(strCitationItem) = True) Then
            'gets the matches (the year of issue)
            Set colMatchesYearFromCitationItem = objRegExpYearFromCitationItem.Execute(strCitationItem)

            'treats all matches (there should only one match)
            For Each objMatchYearFromCitationItem In colMatchesYearFromCitationItem
                'gets only the year, without the extra characters in the match
                'from '"issued":{"date-parts":[["Year"' or '"issued":{"date-parts":[[Year' to just 'Year'
                If Right(objMatchYearFromCitationItem.value, 1) = Chr(34) Then
                    'takes only the year and removes existing '"'
                    strYearFromCitationItem = Replace(Right(objMatchYearFromCitationItem.value, 5), Chr(34), "")
                Else
                    strYearFromCitationItem = Right(objMatchYearFromCitationItem.value, 4)
                End If
            Next objMatchYearFromCitationItem 'treats all matches (there should be at most one match, zero when editors are listed instead of the authors)
        End If 'checks that the string can be compared

    End If 'if the citation item's full info is not empty


    'returns the year of issue
    GAUG_getYearFromCitationItem = strYearFromCitationItem

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                                                                        **
'**  Last modified: 2024-10-02                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getPartsFromVisibleCitationItem(ByVal strVisibleCitationItem As String) As Variant()                                       **
'**                                                                                                                                           **
'**  Returns the authors or editors (if present), year of issue and letter (if present) of a particular visible citation item.                **
'**                                                                                                                                           **
'**  Parameter strVisibleCitationItem is a string that contains the visible text of the citation item                                         **
'**                                                                                                                                           **
'**  RETURNS: An array that contains the authors or editors, the year of issue and the letter after the year of the visible citation item.    **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_getPartsFromVisibleCitationItem(ByVal strVisibleCitationItem As String) As Variant()

    Dim varPartsFromVisibleCitationItem(1 To 3) As Variant
    Dim intSizeOfString As Integer


    'initializes the parts to emtpy string
    '(in some cases the authors will not be present or the year of issue does not have a letter at the end)
    'e.g. in a citation such as '(FamilyName, 2024a, 2024b)', the second item in the citation is '2024b' (the first is 'FamilyName, 2024a') which does not include the author but includes a letter
    'e.g. in a citation such as '(FamilyName, 2020; OtherFamilyName, 2023)', both items have authors but not letter at the end of the year
    varPartsFromVisibleCitationItem(1) = ""
    varPartsFromVisibleCitationItem(2) = ""
    varPartsFromVisibleCitationItem(3) = ""


    'removes leading or trailing blank spaces
    strVisibleCitationItem = Trim(strVisibleCitationItem)


    'if the citation item is not empty
    If Not strVisibleCitationItem = "" Then
        'gets the size of the string containing the visible citation item
        intSizeOfString = Len(strVisibleCitationItem)

        'if the citation item DOES NOT include a letter at the end of the year
        If Asc(Right(strVisibleCitationItem, 1)) >= 48 And Asc(Right(strVisibleCitationItem, 1)) <= 57 Then
            'if the visible citation item includes authors (or editors) its length is bigger than ', YYYY'
            If intSizeOfString > 6 Then
                'gets the authors or editors
                varPartsFromVisibleCitationItem(1) = Mid(strVisibleCitationItem, 1, intSizeOfString - 6)
            End If
            'gets the year
            varPartsFromVisibleCitationItem(2) = Mid(strVisibleCitationItem, intSizeOfString - 3, 4)

        'if the citation item includes a letter at the end of the year
        Else
            'if the visible citation item includes authors (or editors) its length is bigger than ', YYYYx'
            If intSizeOfString > 7 Then
                'gets the authors or editors
                varPartsFromVisibleCitationItem(1) = Mid(strVisibleCitationItem, 1, intSizeOfString - 7)
            End If
            'gets the year
            varPartsFromVisibleCitationItem(2) = Mid(strVisibleCitationItem, intSizeOfString - 4, 4)
            'gets the letter at the end of the year
            varPartsFromVisibleCitationItem(3) = Right(strVisibleCitationItem, 1)
        End If 'if the citation item DOES NOT include a letter at the end of the year

    End If


    'returns the three parts of the visible citation item
    GAUG_getPartsFromVisibleCitationItem = varPartsFromVisibleCitationItem

End Function



'*****************************************************************************************
'*****************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                  **
'**  Last modified: 2024-10-07                                                          **
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

    Dim intAvailableMendeleyVersion, intUseMendeleyVersion As Integer
    Dim documentSection As Section
    Dim blnFound, blnBibliographyFound, blnCitationFound, blnReferenceEntryFound, blnCitationEntryFound, blnGenerateHyperlinksForURLs, blnURLFound As Boolean
    Dim intRefereceNumber, i As Integer
    Dim objRegExpBiblioEntries, objRegExpVisibleCitationItems, objRegExpFindHiddenCitationItems, objRegExpFindBiblioEntry, objRegExpFindVisibleCitationItem, objRegExpURL As RegExp
    Dim colMatchesBiblioEntries, colMatchesVisibleCitationItems, colMatchesFindHiddenCitationItems, colMatchesFindBiblioEntry, colMatchesFindVisibleCitationItem, colMatchesURL As MatchCollection
    Dim objMatchBiblioEntry, objMatchVisibleCitationItem, objMatchsFindHiddenCitationItem, objMatchFindBiblioEntry, objMatchURL As match
    Dim strTempMatch, strSubStringOfTempMatch, strLastAuthors As String
    Dim strTypeOfExecution As String
    Dim blnMabEntwickeltSich As Boolean
    Dim stlStyleInDocument As Word.Style
    Dim strStyleForTitleOfBibliography As String
    Dim blnStyleForTitleOfBibliographyFound As Boolean
    Dim strURL As String
    Dim arrNonDetectedURLs, varNonDetectedURL As Variant
    Dim strDoHyperlinksExist As String
    Dim intTotalNumberOfFieldsOrContentControls, intIndexCurrentFieldOrContentControl As Integer
    Dim objCurrentFieldOrContentControl As Object
    Dim strAllCitationsFullInformation, strCitationFullInfo As String
    Dim varCitationItemsFromCitationFullInfo() As Variant
    Dim varPartsFromVisibleCitationItem() As Variant
    Dim varAuthorsFomCitationItem() As Variant
    Dim varEditorsFomCitationItem() As Variant
    Dim varYearFomCitationItem As String
    Dim intAuthorFromCitationItem, intEditorFomCitationItem As Integer
    Dim intCitationItemFromCitationFullInfo As Integer
    Dim strOrphanCitationItems As String
    Dim varFieldsOrContentControls As Variant
    Dim currentPosition As range


'*****************************
'*   Custom configuration    *
'*****************************
    'possible values are 0, 1 or 2
    'SEE DOCUMENTATION
    'set to 0 if the macros should automatically detect the version of Mendeley
    'set to 1 if the macros should use Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic)
    'set to 2 if the macros should use Mendeley Reference Manager 2.x (with the App Mendeley Cite)
    intUseMendeleyVersion = 0

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
    'When Mendeley Reference Manager 2.x is used, ONLY "RemoveHyperlinks" is available
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

    'gets the version of Mendeley (autodetect or specified)
    intAvailableMendeleyVersion = GAUG_getAvailableMendeleyVersion(intUseMendeleyVersion)

    'gets all the hidden information of all the citations (if Mendeley Reference Manager 2.x is used)
    strAllCitationsFullInformation = GAUG_getAllCitationsFullInformation(intAvailableMendeleyVersion)
    
    'gets the current position of the cursor on the document
    Set currentPosition = Selection.range






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
    Set objRegExpBiblioEntries = New RegExp
    'sets the pattern to match every reference entry in the bibliography (it may include a character of carriage return)
    '(all text from the beginning of the string, or carriage return, until a year between parentheses is found)
    'updated to include "(Ed.)" and "(Eds.)" when editors are used for the citations and bibliography
    objRegExpBiblioEntries.Pattern = "((^)|(\r))[^(\r)]*(\(Eds?\.\)\.\s)?\(\d\d\d\d[a-zA-Z]?\)"
    'sets case insensitivity
    objRegExpBiblioEntries.IgnoreCase = False
    'sets global applicability
    objRegExpBiblioEntries.Global = True
    'creates the object for regular expressions (to get all URLs in biblio)
    Set objRegExpURL = New RegExp
    'sets the pattern to match every URL in the bibliography (http, https or ftp)
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
            'according to the version of Mendeley
            Select Case intAvailableMendeleyVersion
                'Mendeley Desktop 1.x is installed
                Case 1
                        'gets the list of fields in this section
                        Set varFieldsOrContentControls = documentSection.range.Fields
                'Mendeley Reference Manager 2.x is installed
                Case 2
                        'gets the list of content controls in this section
                        Set varFieldsOrContentControls = documentSection.range.ContentControls
            End Select

            'checks all fields or content controls
            For Each objCurrentFieldOrContentControl In varFieldsOrContentControls

                'according to the version of Mendeley
                Select Case intAvailableMendeleyVersion
                    'Mendeley Desktop 1.x is installed
                    Case 1
                            'checks if it is the bibliography
                            If objCurrentFieldOrContentControl.Type = wdFieldAddin And Trim(objCurrentFieldOrContentControl.Code) = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY" Then
                                blnBibliographyFound = True
                            End If
                    'Mendeley Reference Manager 2.x is installed
                    Case 2
                            'checks if it is the bibliograpy
                            If objCurrentFieldOrContentControl.Type = wdContentControlRichText And Trim(objCurrentFieldOrContentControl.Tag) = "MENDELEY_BIBLIOGRAPHY" Then
                                blnBibliographyFound = True
                            End If
                End Select


                'if it is the bibliography
                If blnBibliographyFound Then
                    'start the numbering
                    intRefereceNumber = 1

                    'according to the version of Mendeley
                    Select Case intAvailableMendeleyVersion
                        'Mendeley Desktop 1.x is installed
                        Case 1
                                'selects the current field (Mendeley's bibliography field)
                                objCurrentFieldOrContentControl.Select
                        'Mendeley Reference Manager 2.x is installed
                        Case 2
                                'selects the current content control (Mendeley's bibliography content control)
                                objCurrentFieldOrContentControl.range.Select
                    End Select

                    'checks that the string can be compared
                    If (objRegExpBiblioEntries.Test(Selection) = True) Then
                        'gets the matches (all entries in bibliography according to the regular expression)
                        Set colMatchesBiblioEntries = objRegExpBiblioEntries.Execute(Selection)

                        'treats all matches (all entries in bibliography) to generate bookmars
                        '(we have to find AGAIN every entry to select it and create the bookmark)
                        For Each objMatchBiblioEntry In colMatchesBiblioEntries
                            'removes the carriage return from match, if necessary
                            strTempMatch = Replace(objMatchBiblioEntry.value, Chr(13), "")

                            'prevents errors if the match is longer than 256 characters
                            'however, the reference will not be linked if it has more than 20 authors (APA 7th edition)
                            'or more than 7 authors (APA 6th edition) due to the fact that APA replaces some of them with ellipsis (...)
                            strSubStringOfTempMatch = Left(strTempMatch, 256)

                            'according to the version of Mendeley
                            Select Case intAvailableMendeleyVersion
                                'Mendeley Desktop 1.x is installed
                                Case 1
                                        'selects the current field (Mendeley's bibliography field)
                                        objCurrentFieldOrContentControl.Select
                                'Mendeley Reference Manager 2.x is installed
                                Case 2
                                        'selects the current content control (Mendeley's bibliography content control)
                                        objCurrentFieldOrContentControl.range.Select
                            End Select

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
                                    Name:="GAUG_SignetBibliographie_" & format(CStr(intRefereceNumber), "00#"), _
                                    range:=Selection.range
                            End If

                            'continues with the next number
                            intRefereceNumber = intRefereceNumber + 1

                        Next 'treats all matches (all entries in bibliography) to generate bookmars
                    End If

                    'by now, we have created all bookmarks and have all entries in colMatchesBiblioEntries
                    'for future use when creating the hyperlinks

                    'generates the hyperlinks for the URLs in the bibliography, if required
                    If blnGenerateHyperlinksForURLs Then

                        'generates the hyperlnks from the list of non detected URLs
                        'the non detected URLs shall be done first or some conflicts may arise
                        For Each varNonDetectedURL In arrNonDetectedURLs
                            'according to the version of Mendeley
                            Select Case intAvailableMendeleyVersion
                                'Mendeley Desktop 1.x is installed
                                Case 1
                                        'selects the current field (Mendeley's bibliography field)
                                        objCurrentFieldOrContentControl.Select
                                'Mendeley Reference Manager 2.x is installed
                                Case 2
                                        'selects the current content control (Mendeley's bibliography content control)
                                        objCurrentFieldOrContentControl.range.Select
                            End Select

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

                        'according to the version of Mendeley
                        Select Case intAvailableMendeleyVersion
                            'Mendeley Desktop 1.x is installed
                            Case 1
                                    'selects the current field (Mendeley's bibliography field)
                                    objCurrentFieldOrContentControl.Select
                            'Mendeley Reference Manager 2.x is installed
                            Case 2
                                    'selects the current content control (Mendeley's bibliography content control)
                                    objCurrentFieldOrContentControl.range.Select
                        End Select

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

                                'according to the version of Mendeley
                                Select Case intAvailableMendeleyVersion
                                    'Mendeley Desktop 1.x is installed
                                    Case 1
                                            'selects the current field (Mendeley's bibliography field)
                                            objCurrentFieldOrContentControl.Select
                                    'Mendeley Reference Manager 2.x is installed
                                    Case 2
                                            'selects the current content control (Mendeley's bibliography content control)
                                            objCurrentFieldOrContentControl.range.Select
                                End Select

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

                    'exits the for loop, the bibliography has ben found already
                    Exit For

                End If 'if it is the biblio
            Next 'checks all fields or content controls
        End If 'checks if the bibliography is in this section

        'if the bibliography has been found already, no need to check other sections
        If blnBibliographyFound Then
            'exits the for loop, the bibliography has ben found already
            Exit For
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
    Set objRegExpVisibleCitationItems = New RegExp
    Set objRegExpFindHiddenCitationItems = New RegExp
    Set objRegExpFindBiblioEntry = New RegExp
    Set objRegExpFindVisibleCitationItem = New RegExp
    'sets the pattern to match every citation entry (with or without authors) in current field
    '(the year of publication is always present, authors may not be present)
    '(all text non starting by "(" or "," or ";" followed by non digits until a year is found)
    objRegExpVisibleCitationItems.Pattern = "[^(\(|,|;)][^0-9]*\d\d\d\d[a-zA-Z]?"
    'sets the pattern to match every citation entry from the data of the current field
    'original regular expression to get the authors info from Field.Code "((\"id\")|(\"family\")|(\"given\"))\s\:\s\"[^\"]*\""
    '(all text related to "id", "family" and "given"), all '\"' were substituted by '\" & Chr(34) & "'
    'objRegExpFindHiddenCitationItems.Pattern = "((\" & Chr(34) & "id\" & Chr(34) & ")|(\" & Chr(34) & "family\" & Chr(34) & ")|(\" & Chr(34) & "given\" & Chr(34) & "))\s\:\s\" & Chr(34) & "[^\" & Chr(34) & "]*\" & Chr(34)
    'updated to separate authors from editors:
    'objRegExpFindHiddenCitationItems.Pattern = "(\" & Chr(34) & "editor\" & Chr(34) & "\s:\s)|(((\" & Chr(34) & "id\" & Chr(34) & ")|(\" & Chr(34) & "family\" & Chr(34) & "))\s\:\s\" & Chr(34) & "[^\" & Chr(34) & "]*\" & Chr(34) & ")"
    'updated to also include the publication year:
    objRegExpFindHiddenCitationItems.Pattern = "((\" & Chr(34) & "editor\" & Chr(34) & "\s*:\s*)|(((\" & Chr(34) & "id\" & Chr(34) & ")|(\" & Chr(34) & "family\" & Chr(34) & "))\s*\:\s*\" & Chr(34) & "[^\" & Chr(34) & "]*\" & Chr(34) & "))|(\[\s*\[\s*\" & Chr(34) & "[0-9]+\" & Chr(34) & "\s*\]\s*\])|(\[\s*\[\s*\" & Chr(34) & "[0-9]+\" & Chr(34) & "\s*,\s*\" & Chr(34) & "[0-9]+\" & Chr(34) & "\s*\]\s*\])"
    'sets case insensitivity
    objRegExpVisibleCitationItems.IgnoreCase = False
    objRegExpFindHiddenCitationItems.IgnoreCase = False
    objRegExpFindBiblioEntry.IgnoreCase = False
    objRegExpFindVisibleCitationItem.IgnoreCase = False
    'sets global applicability
    objRegExpVisibleCitationItems.Global = True
    objRegExpFindHiddenCitationItems.Global = True
    objRegExpFindBiblioEntry.Global = True
    objRegExpFindVisibleCitationItem.Global = True
    strOrphanCitationItems = ""

    
    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        
        'according to the version of Mendeley
        Select Case intAvailableMendeleyVersion
            'Mendeley Desktop 1.x is installed
            Case 1
                    'counts the number of fields in this section (needed to iterate over all of them)
                    'intTotalNumberOfFieldsOrContentControls = documentSection.range.Fields.Count
                    Set varFieldsOrContentControls = documentSection.range.Fields
            'Mendeley Reference Manager 2.x is installed
            Case 2
                    'counts the number of content controls in this section (needed to iterate over all of them)
                    'intTotalNumberOfFieldsOrContentControls = documentSection.range.ContentControls.Count
                    Set varFieldsOrContentControls = documentSection.range.ContentControls
        End Select

        'checks all fields or content controls
        For Each objCurrentFieldOrContentControl In varFieldsOrContentControls
        'For intIndexCurrentFieldOrContentControl = 1 To intTotalNumberOfFieldsOrContentControls

            'according to the version of Mendeley
            Select Case intAvailableMendeleyVersion
                'Mendeley Desktop 1.x is installed
                Case 1
                        'gets the current field in this document section
                        'Set objCurrentFieldOrContentControl = documentSection.range.Fields(intIndexCurrentFieldOrContentControl)
                        'checks if it is a citation
                        If objCurrentFieldOrContentControl.Type = wdFieldAddin And Left(objCurrentFieldOrContentControl.Code, 18) = "ADDIN CSL_CITATION" Then
                            blnCitationFound = True
                        Else
                            blnCitationFound = False
                        End If
                'Mendeley Reference Manager 2.x is installed
                Case 2
                        'gets the current content control in this document section
                        'Set objCurrentFieldOrContentControl = documentSection.range.ContentControls(intIndexCurrentFieldOrContentControl)
                        'checks if it is a citation
                        If objCurrentFieldOrContentControl.Type = wdContentControlRichText And Left(Trim(objCurrentFieldOrContentControl.Tag), 21) = "MENDELEY_CITATION_v3_" Then
                            blnCitationFound = True
                        Else
                            blnCitationFound = False
                        End If
            End Select

            'if it is a citation
            If blnCitationFound Then

                'according to the version of Mendeley
                Select Case intAvailableMendeleyVersion
                    'Mendeley Desktop 1.x is installed
                    Case 1
                            'selects the current field (Mendeley's citation field)
                            objCurrentFieldOrContentControl.Select
                            'gets the full information of the current citation (which contains all individual citations in the field or content control object with the full list of authors)
                            strCitationFullInfo = GAUG_getCitationFullInfo(intAvailableMendeleyVersion, strAllCitationsFullInformation, objCurrentFieldOrContentControl, Nothing)
                            'gets the list of all items (individual citations within the field or content control) from the citation full information
                            varCitationItemsFromCitationFullInfo = GAUG_getCitationItemsFromCitationFullInfo(intAvailableMendeleyVersion, strCitationFullInfo)
                    
                    'Mendeley Reference Manager 2.x is installed
                    Case 2
                            'selects the current content control (Mendeley's citation content control)
                            objCurrentFieldOrContentControl.range.Select
                            'gets the full information of the current citation (which contains all individual citations in the field or content control object with the full list of authors)
                            strCitationFullInfo = GAUG_getCitationFullInfo(intAvailableMendeleyVersion, strAllCitationsFullInformation, Nothing, objCurrentFieldOrContentControl)
                            'gets the list of all items (individual citations within the field or content control) from the citation full information
                            varCitationItemsFromCitationFullInfo = GAUG_getCitationItemsFromCitationFullInfo(intAvailableMendeleyVersion, strCitationFullInfo)
                End Select


                'checks that the string can be compared (both, Selection and the string with the full information of the citation)
                If (objRegExpVisibleCitationItems.Test(Selection) = True) And (objRegExpFindHiddenCitationItems.Test(strCitationFullInfo) = True) Then
                    'gets the matches (all entries in the citation according to the regular expression)
                    Set colMatchesVisibleCitationItems = objRegExpVisibleCitationItems.Execute(Selection)
                    'gets the matches (all entries in the full information of the citation according to the regular expression)
                    '(used to find the entry in the bibliography)
                    Set colMatchesFindHiddenCitationItems = objRegExpFindHiddenCitationItems.Execute(strCitationFullInfo)

                    'treats all matches (all entries in citation) to generate hyperlinks
                    For Each objMatchVisibleCitationItem In colMatchesVisibleCitationItems
                        'I COULD NOT FIND A MORE EFFICIENT WAY TO SELECT EVERY REFERENCE
                        'IN ORDER TO CREATE THE LINK:
                        'Start: Needs re-work

                        'gets the list of all parts (author, year, and letter of year if present) from the visible citation item (entry in visible text of citation)
                        'position 1 is the author (when present)
                        'position 2 is the issue year
                        'position 3 is the letter after the issue year (when present)
                        varPartsFromVisibleCitationItem = GAUG_getPartsFromVisibleCitationItem(objMatchVisibleCitationItem.value)


                        'when citations are merged, they are ordered by the authors' family names
                        'the position of the citation in the visible text may not correspond to the position in the citation hidden data,
                        'we need to find the entry, but we may not have the authors's family names :(


                        'if the current match has authors's family names (not only the year)
                        'we keep them stored for future use if next citation item DOES NOT include them
                        If Len(varPartsFromVisibleCitationItem(1)) > 0 Then
                            strLastAuthors = varPartsFromVisibleCitationItem(1)
                        End If

                        'checks if the list of all items from the citation full information is not empty (see https://riptutorial.com/excel-vba/example/30824/check-if-array-is-initialized--if-it-contains-elements-or-not--)
                        If Not Not varCitationItemsFromCitationFullInfo Then
                            'treats all citation items from the citation full information (to find which one corresponds to the current visible citation item being treated)
                            For intCitationItemFromCitationFullInfo = 1 To UBound(varCitationItemsFromCitationFullInfo)
                                'gets the list of authors (if available) from the citation item
                                varAuthorsFomCitationItem = GAUG_getAuthorsEditorsFromCitationItem(intAvailableMendeleyVersion, "author", varCitationItemsFromCitationFullInfo(intCitationItemFromCitationFullInfo))
                                'gets the list of editors (if available) from the citation item
                                varEditorsFomCitationItem = GAUG_getAuthorsEditorsFromCitationItem(intAvailableMendeleyVersion, "editor", varCitationItemsFromCitationFullInfo(intCitationItemFromCitationFullInfo))
                                'gets the year of issue from the citation item
                                varYearFomCitationItem = GAUG_getYearFromCitationItem(intAvailableMendeleyVersion, varCitationItemsFromCitationFullInfo(intCitationItemFromCitationFullInfo))


                                'initializes the regular expressions
                                objRegExpFindBiblioEntry.Pattern = ""
                                objRegExpFindVisibleCitationItem.Pattern = ""


                                'if the citation item has authors (instead of editors)
                                If Not Not varAuthorsFomCitationItem Then
                                    For intAuthorFromCitationItem = 1 To UBound(varAuthorsFomCitationItem)
                                        'gets the last name of the author and adds it to the regular expression (used to find the entry in the bibliography)
                                        objRegExpFindBiblioEntry.Pattern = objRegExpFindBiblioEntry.Pattern & varAuthorsFomCitationItem(intAuthorFromCitationItem) & ".*"
                                        'creates another regular expression to match the entry in the bibliography with the citation item in the visible text, they are not in the same position as thought
                                        objRegExpFindVisibleCitationItem.Pattern = objRegExpFindVisibleCitationItem.Pattern & varAuthorsFomCitationItem(intAuthorFromCitationItem) & ".*"
                                        'if this is the first authorof many, this could be the only one listed, and the rest as "et al."
                                        If intAuthorFromCitationItem = 1 And UBound(varAuthorsFomCitationItem) > 1 Then
                                            'includes the part to check for "et al." (only for the visible citation item, the entry in the bibliography has the full list)
                                            objRegExpFindVisibleCitationItem.Pattern = objRegExpFindVisibleCitationItem.Pattern & "((et al\..*)|("
                                        End If
                                    Next
                                    'closes the parenthesis in the pattern if more than one author
                                    If UBound(varAuthorsFomCitationItem) > 1 Then
                                        objRegExpFindVisibleCitationItem.Pattern = objRegExpFindVisibleCitationItem.Pattern & "))"
                                    End If

                                'but if no authors were found (like with a book with only editors), we use editors instead
                                Else
                                    'if the citation item has editors
                                    If Not Not varEditorsFomCitationItem Then
                                        For intEditorFromCitationItem = 1 To UBound(varEditorsFomCitationItem)
                                            'gets the last name of the editor and adds it to the regular expression (used to find the entry in the bibliography)
                                            objRegExpFindBiblioEntry.Pattern = objRegExpFindBiblioEntry.Pattern & varEditorsFomCitationItem(intEditorFromCitationItem) & ".*"
                                            'creates another regular expression to match the entry in the bibliography with the citation item in the visible text, they are not in the same position as thought
                                            objRegExpFindVisibleCitationItem.Pattern = objRegExpFindVisibleCitationItem.Pattern & varEditorsFomCitationItem(intEditorFromCitationItem) & ".*"
                                            'if this is the first editor of many, this could be the only one listed, and the rest as "et al."
                                            If intEditorFromCitationItem = 1 And UBound(varEditorsFomCitationItem) > 1 Then
                                                'includes the part to check for "et al." (only for the visible citation item, the entry in the bibliography has the full list)
                                                objRegExpFindVisibleCitationItem.Pattern = objRegExpFindVisibleCitationItem.Pattern & "((et al\..*)|("
                                            End If
                                        Next
                                        'closes the parenthesis in the pattern if more than one editor
                                        If UBound(varEditorsFomCitationItem) > 1 Then
                                            objRegExpFindVisibleCitationItem.Pattern = objRegExpFindVisibleCitationItem.Pattern & "))"
                                        End If
                                    End If
                                End If

                                'finishes the patterns including the year and the letter shown in the visible citation item
                                objRegExpFindVisibleCitationItem.Pattern = objRegExpFindVisibleCitationItem.Pattern & varYearFomCitationItem & varPartsFromVisibleCitationItem(3)
                                objRegExpFindBiblioEntry.Pattern = objRegExpFindBiblioEntry.Pattern & "\(" & varYearFomCitationItem & varPartsFromVisibleCitationItem(3) & "\)"
                                'MsgBox objMatchVisibleCitationItem.value & " -> *" & _
                                '    varPartsFromVisibleCitationItem(1) & "*" & varPartsFromVisibleCitationItem(2) & "*" & varPartsFromVisibleCitationItem(3) & "*" & vbCrLf & vbCrLf & _
                                '    objRegExpFindVisibleCitationItem.Pattern & vbCrLf & vbCrLf & _
                                '    objRegExpFindBiblioEntry.Pattern


                                'if the current visible citation item has authors's family names (not only the year)
                                If Len(varPartsFromVisibleCitationItem(1)) > 0 Then
                                    'checks if this item from the citation full information corresponds to the visible citation item being treated
                                    Set colMatchesFindVisibleCitationItem = objRegExpFindVisibleCitationItem.Execute(objMatchVisibleCitationItem.value)
                                Else
                                    'checks if this item from the citation full information corresponds to the visible citation item being treated
                                    Set colMatchesFindVisibleCitationItem = objRegExpFindVisibleCitationItem.Execute(strLastAuthors & ", " & objMatchVisibleCitationItem.value)
                                End If
                                
                                'if this item from the citation full information corresponds to the visible citation item being treated
                                If colMatchesFindVisibleCitationItem.Count > 0 Then
                                    'MsgBox ("Match between DOCUMENT (visible citation item) and DATA (hidden citation item) found:" & vbCrLf & vbCrLf & _
                                    '    colMatchesFindVisibleCitationItem.Item(0).value)
                                    Exit For
                                End If

                            Next 'treats all citation items from the citation full information
                        End If 'checks if the list of all items from the citation full information is not empty



                        'last verification to make sure we are here because this item from the citation full information corresponds
                        'to the visible citation item being treated and not because the for loop reached the end
                        If colMatchesFindVisibleCitationItem.Count = 0 Then
                            'cleans the regular expression as no matches were found
                            objRegExpFindBiblioEntry.Pattern = "Error: Citation not found"
                        End If

                        'at this point, the regular expression to find the entry in the bibliography is ready
                        'MsgBox "The visible citation item" & vbCrLf & _
                        '    objMatchVisibleCitationItem.value & vbCrLf & _
                        '    "matches:" & vbCrLf & vbCrLf & _
                        '    objRegExpFindVisibleCitationItem.Pattern & vbCrLf & _
                        '    objRegExpFindBiblioEntry.Pattern


                        'it is time to find the citation entry in the bibliography and link the current visible citation item


                        'initializes the position
                        i = 1
                        'finds the position of the citation entry in the list of references in the bibliography
                        blnReferenceEntryFound = False
                        For Each objMatchBiblioEntry In colMatchesBiblioEntries
                            'MsgBox ("Searching for citation in bibliography:" & vbCrLf & vbCrLf & "Using..." & vbCrLf & objRegExpFindBiblioEntry.Pattern & vbCrLf & objMatchBiblioEntry.value)
                            'gets the matches, if any, to check if this reference entry corresponds to the visible citation item being treated
                            Set colMatchesFindBiblioEntry = objRegExpFindBiblioEntry.Execute(objMatchBiblioEntry.value)
                            'if the this is the corresponding reference entry
                            'Verify for MabEntwickeltSich: perhaps a more strict verification is needed
                            If colMatchesFindBiblioEntry.Count > 0 Then
                                blnReferenceEntryFound = True
                                Exit For
                            End If
                            'continues with the next number
                            i = i + 1
                        Next

                        'at this point we also have the position (i) in the biblio, we are ready to create the hyperlink
                        'the position is isued to link to the bookmark 'GAUG_SignetBibliographie_<position>'

                        'if reference entry was found (shall always find it), creates the hyperlink
                        If blnReferenceEntryFound Then
                            'MsgBox ("Citation was found in the bibliography" & vbCrLf & vbCrLf & colMatchesFindBiblioEntry.Item(0).value)
                            'according to the version of Mendeley
                            Select Case intAvailableMendeleyVersion
                                'Mendeley Desktop 1.x is installed
                                Case 1
                                        'selects the current field (Mendeley's citation field)
                                        objCurrentFieldOrContentControl.Select
                                'Mendeley Reference Manager 2.x is installed
                                Case 2
                                        'selects the current content control (Mendeley's citation content control)
                                        objCurrentFieldOrContentControl.range.Select
                            End Select

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
                                Selection.MoveEnd Unit:=wdCharacter, Count:=objMatchVisibleCitationItem.FirstIndex + objMatchVisibleCitationItem.Length - 1
                                'if the first character is a blank space
                                If Left(objMatchVisibleCitationItem.value, 1) = " " Then
                                    'removes the leading blank space
                                    Selection.MoveStart Unit:=wdCharacter, Count:=objMatchVisibleCitationItem.FirstIndex + 1
                                Else
                                    'uses the whole range
                                    Selection.MoveStart Unit:=wdCharacter, Count:=objMatchVisibleCitationItem.FirstIndex
                                End If

                                'creates the hyperlink for the current citation entry
                                'a cross-reference is not a good idea, it changes the text in citation (or may delete citation):
                                'Selection.Fields.Add Range:=Selection.Range, _
                                '    Type:=wdFieldEmpty, _
                                '    Text:="REF " & Chr(34) & "GAUG_SignetBibliographie_" & Format(CStr(i), "00#") & Chr(34) & " \h", _
                                '    PreserveFormatting:=True
                                'better to use normal hyperlink:
                                Selection.Hyperlinks.Add Anchor:=Selection.range, _
                                    Address:="", SubAddress:="GAUG_SignetBibliographie_" & format(CStr(i), "00#"), _
                                    ScreenTip:=""

                            End If
                        Else
                            'if the visible citation item could not be linked to an entry in the bibliography
                            strOrphanCitationItems = strOrphanCitationItems & Trim(objMatchVisibleCitationItem.value) & vbCrLf
                        End If

                        'Ends: Needs re-work

                        'at this point current citation entry is linked to corresponding reference in biblio

                    Next 'treats all matches (all entries in citation) to generate hyperlinks

                End If 'checks that the string can be compared

            End If 'if it is a citation
        Next 'checks all fields or content controls

        'at this point all citations are linked to their corresponding reference in biblio

    Next 'documentSection

    'if orphan citations exist
    If Len(strOrphanCitationItems) > 0 Then
        MsgBox "Orphan citation entries found:" & vbCrLf & vbCrLf & _
            strOrphanCitationItems & vbCrLf & _
            "Remove them from document!", _
            vbExclamation, "GAUG_createHyperlinksForCitationsAPA()"
    End If

    'returns to original position in the document
    currentPosition.Select
    
    'reenables the screen updating
    Application.ScreenUpdating = True

End Sub



'*****************************************************************************************
'*****************************************************************************************
'**  Author: Jos� Luis Gonz�lez Garc�a                                                  **
'**  Last modified: 2024-10-05                                                          **
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
    Dim objRegExpVisibleCitationItems, objRegExpURL As RegExp
    Dim colMatchesVisibleCitationItems, colMatchesURL As MatchCollection
    Dim objMatchVisibleCitationItem, objMatchURL As match
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
            "You cannot call the macro GAUG_removeHyperlinksForCitations(strTypeOfExecution)" & vbCrLf & _
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
                                Name:="GAUG_SignetBibliographie_" & format(CStr(intRefereceNumber), "00#"), _
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
    Set objRegExpVisibleCitationItems = New RegExp
    'sets the pattern to match every citation entry in current field
    'it should be "[" + Number + "]"
    objRegExpVisibleCitationItems.Pattern = "\[[0-9]+\]"
    'sets case insensitivity
    objRegExpVisibleCitationItems.IgnoreCase = False
    'sets global applicability
    objRegExpVisibleCitationItems.Global = True

    'checks all sections
    For Each documentSection In ActiveDocument.Sections
        'checks all fields
        For Each sectionField In documentSection.range.Fields
            'if it is a citation
            If sectionField.Type = wdFieldAddin And Left(sectionField.Code, 18) = "ADDIN CSL_CITATION" Then

                'selects the current field (Mendeley's citation field)
                sectionField.Select

                'checks that the string can be compared (both, Selection and Field.Code)
                If objRegExpVisibleCitationItems.Test(Selection) = True Then
                    'gets the matches (all entries in the citation according to the regular expression)
                    Set colMatchesVisibleCitationItems = objRegExpVisibleCitationItems.Execute(Selection)

                    'treats all matches (all entries in citation) to generate hyperlinks
                    For Each objMatchVisibleCitationItem In colMatchesVisibleCitationItems
                        'gets the citation number as integer
                        'this will also eliminate leading zeros in numbers (in case of manual modifications)
                        intCitationNumber = CInt(Mid(objMatchVisibleCitationItem.value, 2, Len(objMatchVisibleCitationItem.value) - 2))

                        'to make sure the citation number as text is the same as numeric
                        'and that the citation number is in the bibliography
                        If (("[" & CStr(intCitationNumber) & "]") = objMatchVisibleCitationItem.value) And (intCitationNumber > 0 And intCitationNumber < intRefereceNumber) Then
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
                            '    Text:="REF " & Chr(34) & "GAUG_SignetBibliographie_" & Format(CStr(intCitationNumber), "00#") & Chr(34) & " \h", _
                            '    PreserveFormatting:=True
                            'better to use normal hyperlink:
                            Selection.Hyperlinks.Add Anchor:=Selection.range, _
                                Address:="", SubAddress:="GAUG_SignetBibliographie_" & format(CStr(intCitationNumber), "00#"), _
                                ScreenTip:=""
                        Else
                            MsgBox "Orphan citation entry found:" & vbCrLf & vbCrLf & _
                                objMatchVisibleCitationItem.value & vbCrLf & vbCrLf & _
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
'**  Last modified: 2024-10-05                                                          **
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

    Dim intAvailableMendeleyVersion, intUseMendeleyVersion As Integer
    Dim documentSection As Section
    Dim objCurrentFieldOrContentControl As Object
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
    Dim varFieldsOrContentControls As Variant
    Dim currentPosition As range


'*****************************
'*   Custom configuration    *
'*****************************
    'possible values are 0, 1 or 2
    'SEE DOCUMENTATION
    'set to 0 if the macros should automatically detect the version of Mendeley
    'set to 1 if the macros should use Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic)
    'set to 2 if the macros should use Mendeley Reference Manager 2.x (with the App Mendeley Cite)
    intUseMendeleyVersion = 0

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
            vbCritical, "GAUG_removeHyperlinksForCitations(strTypeOfExecution)"

        'stops the execution
        End
    End If

    'gets the version of Mendeley (autodetect or specified)
    intAvailableMendeleyVersion = GAUG_getAvailableMendeleyVersion(intUseMendeleyVersion)

    'selects the type of execution
    Select Case strTypeOfExecution
        Case "RemoveHyperlinks"
            'nothing to do here
        Case "CleanEnvironment"
            'only available when Mendeley Desktop 1.x is used
            If Not intAvailableMendeleyVersion = 1 Then
                MsgBox "Incompatible execution type " & Chr(34) & strTypeOfExecution & Chr(34) & " for GAUG_removeHyperlinksForCitations(strTypeOfExecution)." & vbCrLf & vbCrLf & _
                    "Only " & Chr(34) & "RemoveHyperlinks" & Chr(34) & " can be used with Mendeley Reference Manager 2.x (with the App Mendeley Cite)." & vbCrLf & vbCrLf & _
                    "Cannot continue creating hyperlinks.", _
                    vbCritical, "GAUG_removeHyperlinksForCitations(strTypeOfExecution)"
                'the execution option is not correct
                End
            End If
            'get the API Client from Mendeley
            Set objMendeleyApiClient = Application.Run("Mendeley.mendeleyApiClient") 'MabEntwickeltSich: This is the way to call the macro directly from Mendeley
        Case "CleanFullEnvironment"
            'only available when Mendeley Desktop 1.x is used
            If Not intAvailableMendeleyVersion = 1 Then
                MsgBox "Incompatible execution type " & Chr(34) & strTypeOfExecution & Chr(34) & " for GAUG_removeHyperlinksForCitations(strTypeOfExecution)." & vbCrLf & vbCrLf & _
                    "Only " & Chr(34) & "RemoveHyperlinks" & Chr(34) & " can be used with Mendeley Reference Manager 2.x (with the App Mendeley Cite)." & vbCrLf & vbCrLf & _
                    "Cannot continue creating hyperlinks.", _
                    vbCritical, "GAUG_removeHyperlinksForCitations(strTypeOfExecution)"
                'the execution option is not correct
                End
            End If
            'gets the Undo Edit Button
            Set cbbUndoEditButton = Application.Run("MendeleyLib.getUndoEditButton") 'MabEntwickeltSich: This is the way to call the macro directly from Mendeley
        Case Else
            'reenables the screen updating
            Application.ScreenUpdating = True

            MsgBox "Unknown execution type " & Chr(34) & strTypeOfExecution & Chr(34) & " for GAUG_removeHyperlinksForCitations(strTypeOfExecution)." & vbCrLf & vbCrLf & _
                "Please, correct the error and try again.", _
                vbCritical, "GAUG_removeHyperlinksForCitations(strTypeOfExecution)"

            'the execution option is not correct
            End
    End Select

    'gets the current position of the cursor on the document
    Set currentPosition = Selection.range

    'disables the screen updating while removing the hyperlinks to increase speed
    Application.ScreenUpdating = False






'*****************************
'*  Cleaning old hyperlinks  *
'*****************************
    'checks all sections
    For Each documentSection In ActiveDocument.Sections

        'according to the version of Mendeley
        Select Case intAvailableMendeleyVersion
            'Mendeley Desktop 1.x is installed
            Case 1
                    'gets the list of fields in this section
                    Set varFieldsOrContentControls = documentSection.range.Fields
            'Mendeley Reference Manager 2.x is installed
            Case 2
                    'gets the list of content controls in this section
                    Set varFieldsOrContentControls = documentSection.range.ContentControls
        End Select

        'checks all fields or content controls
        For Each objCurrentFieldOrContentControl In varFieldsOrContentControls
            'initializes the flag
            blnCitationFound = False

            'according to the version of Mendeley
            Select Case intAvailableMendeleyVersion
                'Mendeley Desktop 1.x is installed
                Case 1
                        'checks if it is a citation
                        If objCurrentFieldOrContentControl.Type = wdFieldAddin And Left(objCurrentFieldOrContentControl.Code, 18) = "ADDIN CSL_CITATION" Then
                            blnCitationFound = True
                        End If
                'Mendeley Reference Manager 2.x is installed
                Case 2
                        'checks if it is a citation
                        If objCurrentFieldOrContentControl.Type = wdContentControlRichText And Left(Trim(objCurrentFieldOrContentControl.Tag), 21) = "MENDELEY_CITATION_v3_" Then
                            blnCitationFound = True
                        End If
            End Select

            'if it is a citation
            If blnCitationFound Then

                'according to the version of Mendeley
                Select Case intAvailableMendeleyVersion
                    'Mendeley Desktop 1.x is installed
                    Case 1
                            'selects the current field (Mendeley's citation field)
                            objCurrentFieldOrContentControl.Select
                    'Mendeley Reference Manager 2.x is installed
                    Case 2
                            'selects the current content control (Mendeley's citation content control)
                            objCurrentFieldOrContentControl.range.Select
                End Select


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
                                    "You cannot call the macro GAUG_removeHyperlinksForCitations(strTypeOfExecution)" & vbCrLf & _
                                    "with argument " & Chr(34) & "RemoveHyperlinks" & Chr(34) & "." & vbCrLf & _
                                    "Use " & Chr(34) & "CleanEnvironment" & Chr(34) & " or " & _
                                    Chr(34) & "CleanFullEnvironment" & Chr(34) & " instead." & vbCrLf & _
                                    "You can also call the respective wrapper function." & vbCrLf & vbCrLf & _
                                    "Cannot continue removing hyperlinks.", _
                                    vbCritical, "GAUG_removeHyperlinksForCitations(strTypeOfExecution)"

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
                        sectionFieldName = Application.Run("ZoteroLib.getMarkName", objCurrentFieldOrContentControl)
                        sectionFieldNewName = objMendeleyApiClient.undoManualFormat(sectionFieldName)
                        Call Application.Run("ZoteroLib.fnRenameMark", objCurrentFieldOrContentControl, sectionFieldNewName) 'MabEntwickeltSich: This is another way to call the macro directly from Mendeley
                        Call Application.Run("ZoteroLib.subSetMarkText", objCurrentFieldOrContentControl, INSERT_CITATION_TEXT) 'MabEntwickeltSich: This is another way to call the macro directly from Mendeley
                    Case "CleanFullEnvironment"
                        'restores the citations to the original state (deletes hyperlinks)
                        'slow version
                        cbbUndoEditButton.Execute
                    End Select

            End If 'if it is a citation
        Next objCurrentFieldOrContentControl 'checks all fields or content controls

    Next 'checks all sections






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
            'according to the version of Mendeley
            Select Case intAvailableMendeleyVersion
                'Mendeley Desktop 1.x is installed
                Case 1
                        'gets the list of fields in this section
                        Set varFieldsOrContentControls = documentSection.range.Fields
                'Mendeley Reference Manager 2.x is installed
                Case 2
                        'gets the list of content controls in this section
                        Set varFieldsOrContentControls = documentSection.range.ContentControls
            End Select

            'checks all fields or content controls
            For Each objCurrentFieldOrContentControl In varFieldsOrContentControls

                'according to the version of Mendeley
                Select Case intAvailableMendeleyVersion
                    'Mendeley Desktop 1.x is installed
                    Case 1
                            'checks if it is the bibliography
                            If objCurrentFieldOrContentControl.Type = wdFieldAddin And Trim(objCurrentFieldOrContentControl.Code) = "ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY" Then
                                blnBibliographyFound = True
                            End If
                    'Mendeley Reference Manager 2.x is installed
                    Case 2
                            'checks if it is the bibliograpy
                            If objCurrentFieldOrContentControl.Type = wdContentControlRichText And Trim(objCurrentFieldOrContentControl.Tag) = "MENDELEY_BIBLIOGRAPHY" Then
                                blnBibliographyFound = True
                            End If
                End Select


                'if it is the bibliography
                If blnBibliographyFound Then
                    'according to the version of Mendeley
                    Select Case intAvailableMendeleyVersion
                        'Mendeley Desktop 1.x is installed
                        Case 1
                                'selects the current field (Mendeley's bibliography field)
                                objCurrentFieldOrContentControl.Select
                        'Mendeley Reference Manager 2.x is installed
                        Case 2
                                'selects the current content control (Mendeley's bibliography content control)
                                objCurrentFieldOrContentControl.range.Select
                    End Select

                    'deletes all bookmarks
                    For Each fieldBookmark In Selection.Bookmarks
                        'deletes current bookmark
                        fieldBookmark.Delete
                    Next

                    'according to the version of Mendeley
                    Select Case intAvailableMendeleyVersion
                        'Mendeley Desktop 1.x is installed
                        Case 1
                                'selects the current field (Mendeley's bibliography field)
                                objCurrentFieldOrContentControl.Select
                        'Mendeley Reference Manager 2.x is installed
                        Case 2
                                'selects the current content control (Mendeley's bibliography content control)
                                objCurrentFieldOrContentControl.range.Select
                    End Select
                    
                    'gets all URL hyperlinks from selection
                    Set selectionHyperlinks = Selection.Hyperlinks
                    'deletes all URL hyperlinks
                    'MsgBox "Total number of hyperlinks in biblio: " & CStr(selectionHyperlinks.Count)
                    For i = selectionHyperlinks.Count To 1 Step -1
                        'deletes the current URL hyperlink
                        selectionHyperlinks(1).Delete
                    Next

                    'exits the for loop, the bibliography has ben found already
                    Exit For

                End If 'if it is the bibliography
            Next 'checks all fields or content controls

        End If 'checks if the bibliography is in this section

        'if the bibliography has been found already, no need to check other sections
        If blnBibliographyFound Then
            'exits the for loop, the bibliography has ben found already
            Exit For
        End If

    Next 'checks all sections


    'if the bibliography could not be located in the document
    If Not blnBibliographyFound Then
        MsgBox "The bibliography could not be located in the document." & vbCrLf & vbCrLf & _
            "Make sure that you have inserted the bibliography via the Mendeley's plugin" & vbCrLf & _
            "and that the custom configuration of the GAUG_* macros is correct." & vbCrLf & vbCrLf & _
            "Cannot continue removing hyperlinks.", _
            vbCritical, "GAUG_removeHyperlinksForCitations(strTypeOfExecution)"

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

    'returns to original position in the document
    currentPosition.Select

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



