Attribute VB_Name = "MendeleyMacros"
Option Explicit



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-10                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_isMendeleyCiteOMaticPluginInstalled() As Boolean                                                                           **
'**                                                                                                                                           **
'**  Checks if the MS Word plugin Mendeley Cite-O-Matic is installed in Microsoft Word.                                                       **
'**                                                                                                                                           **
'**  RETURNS: A boolean that indicates if MS Word plugin Mendeley Cite-O-Matic is installed in Microsoft Word.                                **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_isMendeleyCiteOMaticPluginInstalled() As Boolean

    Dim installedAddin As AddIn
    Dim blnFound As Boolean


    'initializes the flag
    blnFound = False
    'checks all AddIns
    For Each installedAddin In Application.AddIns
        'if Mendeley 1.x has been found
        If Left(installedAddin.name, 11) = "Mendeley-1." Then
            blnFound = True
            'exits loop
            Exit For
        End If
    Next

    'returns true if the MS Word plugin Mendeley Cite-O-Matic is installed
    GAUG_isMendeleyCiteOMaticPluginInstalled = blnFound

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-02                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getAvailableMendeleyVersion(Optional ByVal intUseMendeleyVersion As Integer = 0) As Integer                                **
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
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-24                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getMendeleyWebExtensionXMLFileContents()                                                                                   **
'**                                                                                                                                           **
'**  Copies the current .docx file into a temporary folder and renames it to .zip.                                                            **
'**  Extracts the contents of the .zip file and searches for the file 'word\webextensions\webextension<number>.xml' that                      **
'**     corresponds to Mendeley Reference Manager 2.x (with the App Mendeley Cite).                                                           **
'**  Reads the contents of the file 'word\webextensions\webextension<number>.xml'                                                             **
'**                                                                                                                                           **
'**  RETURNS: A string with the contents (if any) of the file 'webextension<number>.xml' that corresponds                                     **
'**     to Mendeley Reference Manager 2.x (with the App Mendeley Cite).                                                                       **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_getMendeleyWebExtensionXMLFileContents()

    Dim adoStream As Object
    Dim strDocumentName As String
    Dim strDocumentPath As String
    Dim strDocumentFullName As String
    Dim strZipDocumentFullName As String
    Dim strTemporaryFolder As String
    Dim strWebExtensionXMLFullName As String
    Dim strCurrentXMLFile As String
    Dim objFileSystem As Object
    Dim objFile As Object
    Dim objXMLFile As Object
    Dim strXMLFileContents As String
    Dim objShell As Object
    Dim objZipAsFolder As Object


    'initializes the string that will hold the the contents of the file 'webextension<number>.xml' that corresponds to Mendeley Reference Manager 2.x (with the App Mendeley Cite)
    strXMLFileContents = ""

    'checks if the document has a file path (if it is saved)
    '(it is not possible to extract its contents if the document is not saved)
    If ActiveDocument.Path <> "" Then
        'gets the name, path and full name of the active document
        strDocumentName = ActiveDocument.name
        strDocumentPath = ActiveDocument.Path
        strDocumentFullName = ActiveDocument.FullName


        'if Mendeley Macros are running on macOS
        #If Mac Then
        '#####On macOS - Start#####

            'MsgBox "On macOS"

            'builds the path for a temporary folder to unzip the document
            strTemporaryFolder = strDocumentPath & "/" & "MendeleyMacros_GAUG_temp"

            'extracts the zip file
            '(AppleScriptTask can allow only one argument to the AppleScript handler, so we send "strDocumentFullName|strTemporaryFolder", later it will be separated)
            If AppleScriptTask("MendeleyMacrosHelper_Mac.scpt", "unzipWordDocument", strDocumentFullName & "|" & strTemporaryFolder) <> "" Then

                'Microsoft Word may have several web extensions installed
                'we need to find the file that corresponds to Mendeley Reference Manager 2.x (with the App Mendeley Cite)

                'gets the first file name in folder 'word/webextensions/'
                strCurrentXMLFile = Dir(strTemporaryFolder & "/word/webextensions/")
                'iterates over all files in folder 'word/webextensions/'
                Do While strCurrentXMLFile <> ""
                    'if the file name starts with 'webextension' and finishes with '.xml'
                    If (Left(strCurrentXMLFile, 12) = "webextension") And (Right(strCurrentXMLFile, 4) = ".xml") Then
                        'creates the path to the 'webextension<number>.xml' file inside the unzipped folder
                        strWebExtensionXMLFullName = strTemporaryFolder & "/word/webextensions/" & strCurrentXMLFile

                        'opens the XML file, reads and closes it
                        strXMLFileContents = AppleScriptTask("MendeleyMacrosHelper_Mac.scpt", "readUTF8XMLFile", strWebExtensionXMLFullName)

                        'if the files contains '<we:property name="MENDELEY_CITATIONS"', it is the one we are looking for
                        If InStr(1, strXMLFileContents, "<we:property name=" & Chr(34) & "MENDELEY_CITATIONS" & Chr(34), vbTextCompare) > 0 Then
                            'stops the search, we have the file
                            Exit Do
                        Else
                            'clears the contents of the string and continues searching for the file
                            strXMLFileContents = ""
                        End If

                    End If 'if the file name starts with 'webextension' and finishes with '.xml'

                    'gets the next file name in folder 'word/webextensions/'
                    strCurrentXMLFile = Dir()

                Loop 'iterates over all files in folder 'word/webextensions/'

                'cleans up the temporary folder
                AppleScriptTask "MendeleyMacrosHelper_Mac.scpt", "deleteTemporaryFolder", strTemporaryFolder

            End If 'extracts the zip file

        '#####On macOS - End#####

        'if Mendeley Macros are running on Windows
        #Else
        '#####On Windows - Start#####

            'MsgBox "On Windows"

            'sets error handling for testing if object ADODB.Stream is available
            On Error Resume Next
            'creates an ADODB.Stream object (used to read the UTF-8 XML files)
            Set adoStream = CreateObject("ADODB.Stream")
            'checks if the object ADODB.Stream was successfully created
            If Err.Number <> 0 Then
                'The reference to 'Microsoft ActiveX Data Objects 6.1 Library' has not been added, object ADODB.Stream is not available
                '(the reference is automatically added, this error just means that object ADODB.Stream was not created)
                MsgBox "Your document exceeds the maximum number of citations." & vbCrLf & vbCrLf & _
                    "ADODB.Stream is necessary in this extreme case." & vbCrLf & _
                    "Enable the ADODB object and try again." & vbCrLf & vbCrLf & _
                    "Cannot continue creating hyperlinks.", _
                    vbCritical, "GAUG_getMendeleyWebExtensionXMLFileContents()"

                'stops the execution
                End
            End If
            'sets back default error handling by VBA
            On Error GoTo 0

            'sets the character set for the stream
            adoStream.Charset = "UTF-8"

            'initializes the file system object
            Set objFileSystem = CreateObject("Scripting.FileSystemObject")


            'builds the path for a temporary folder to unzip the document
            strTemporaryFolder = Environ$("temp") & "\" & "MendeleyMacros_GAUG_temp"

            'cleans up temporary folder if it exists
            If objFileSystem.FolderExists(strTemporaryFolder) Then
                objFileSystem.DeleteFolder strTemporaryFolder, True
            End If

            'creates temporary folder
            objFileSystem.CreateFolder strTemporaryFolder

            'copies the .docx file and renames it to .zip
            strZipDocumentFullName = strTemporaryFolder & "\" & strDocumentName & ".zip"
            objFileSystem.CopyFile strDocumentFullName, strZipDocumentFullName

            'initializes the shell object
            Set objShell = CreateObject("Shell.Application")
            'opens the zip file (which is the .docx file)
            Set objZipAsFolder = objShell.namespace(CVar(strZipDocumentFullName))

            'if the zip file could be opened
            If Not objZipAsFolder Is Nothing Then
                'extracts the zip file (copies the files from the zip to the temporary folder without showing progress)
                objShell.namespace(CVar(strTemporaryFolder)).CopyHere objZipAsFolder.items, 4

                'Microsoft Word may have several web extensions installed
                'we need to find the file that corresponds to Mendeley Reference Manager 2.x (with the App Mendeley Cite)

                'iterates over all files in folder 'word\webextensions\'
                For Each objFile In objFileSystem.GetFolder(CVar(strTemporaryFolder & "\" & "word\webextensions")).Files
                    'if the file name starts with 'webextension' and finishes with '.xml'
                    If (Left(objFile.name, 12) = "webextension") And (Right(objFile.name, 4) = ".xml") Then
                        'creates the path to the 'webextension<number>.xml' file inside the unzipped folder
                        strWebExtensionXMLFullName = strTemporaryFolder & "\word\webextensions\" & objFile.name

                        'checks if the file 'webextension<number>.xml' exists
                        If objFileSystem.FileExists(strWebExtensionXMLFullName) Then
                            'opens the XML file, reads and closes it
                            adoStream.Open
                            adoStream.LoadFromFile strWebExtensionXMLFullName
                            strXMLFileContents = adoStream.ReadText
                            adoStream.Close

                            'if the files contains '<we:property name="MENDELEY_CITATIONS"', it is the one we are looking for
                            If InStr(1, strXMLFileContents, "<we:property name=" & Chr(34) & "MENDELEY_CITATIONS" & Chr(34), vbTextCompare) > 0 Then
                                'stops the search, we have the file
                                Exit For
                            Else
                                'clears the contents of the string and continues searching for the file
                                strXMLFileContents = ""
                            End If
                        End If 'checks if the file 'webextension<number>.xml' exists

                    End If 'if the file name starts with 'webextension' and finishes with '.xml'
                Next objFile 'iterates over all files in folder 'word\webextensions\'

                'cleans up the temporary folder
                objFileSystem.DeleteFolder strTemporaryFolder, True

            End If 'if the zip file could be opened

        '#####On Windows - End#####
        #End If


    End If 'checks if the document has a file path (if it is saved)


    'returns the contents (if any) of the file 'webextension<number>.xml' that corresponds to Mendeley Reference Manager 2.x (with the App Mendeley Cite)
    GAUG_getMendeleyWebExtensionXMLFileContents = strXMLFileContents

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-23                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getAllCitationsFullInformation(intMendeleyVersion As Integer) As String                                                    **
'**                                                                                                                                           **
'**  Finds and returns the full information of all citations, when available.                                                                 **
'**                                                                                                                                           **
'**  Parameter intMendeleyVersion can have two different values:                                                                              **
'**  1:                                                                                                                                       **
'**     Use version 1.x of Mendeley Desktop                                                                                                   **
'**        The function returns an empty string due to the fact that                                                                          **
'**        Mendeley Desktop 1.x stores the information of each citation inside the field of the citation                                      **
'**  2:                                                                                                                                       **
'**     Use version 2.x of Mendeley Reference Manager                                                                                         **
'**        The function returns the information of all citations in a single string                                                           **
'**                                                                                                                                           **
'**  RETURNS: A string that contains all the information of all citations (when Mendeley Reference Manager 2.x is available).                 **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_getAllCitationsFullInformation(ByVal intMendeleyVersion As Integer) As String

    Dim strWordOpenXML As String
    Dim lngStartPosition, lngFirstPossibleEndPosition, lngSecondPossibleEndPosition As Long
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
            'ActiveDocument.WordOpenXML contains everything on the document, including hidden information about the citations added by Mendeley's plugin
            'we need that hidden information to be able to match every citation to the corresponding entry in the bibliography
            '(the information is stored as one block of data within WordOpenXML)

            'sets error handling for testing if the instruction 'ActiveDocument.WordOpenXML' can return the full document
            On Error Resume Next
            'gets the full XML of the document
            strWordOpenXML = ActiveDocument.WordOpenXML
            'checks if 'ActiveDocument.WordOpenXML' was successful
            If Err.Number <> 0 Then
                'if file 'word\webextensions\webextension<number>.xml' is bigger than 1MB, the instruction 'ActiveDocument.WordOpenXML' fails
                'better to directly read the file 'word\webextensions\webextension<number>.xml'
                strWordOpenXML = GAUG_getMendeleyWebExtensionXMLFileContents()
            End If
            'sets back default error handling by VBA
            On Error GoTo 0

            'gets the initial position of the text that contains the full information of all citations
            '(citations start with '<we:property name="MENDELEY_CITATIONS"')
            lngStartPosition = InStr(strWordOpenXML, "<we:property name=" & Chr(34) & "MENDELEY_CITATIONS" & Chr(34))
            'gets the position of the start of Mendeley's citation style (this could be the position where the text that contains the full information of all citations ends)
            '(citation style start with '<we:property name="MENDELEY_CITATIONS_STYLE"')
            lngFirstPossibleEndPosition = InStr(strWordOpenXML, "<we:property name=" & Chr(34) & "MENDELEY_CITATIONS_STYLE" & Chr(34))
            'gets the position of the end of the block of text (this could be the position where the text that contains the full information of all citations ends)
            '(the end of the block starts with '</we:properties><we:bindings/>')
            lngSecondPossibleEndPosition = InStr(strWordOpenXML, "</we:properties><we:bindings/>")


            'if the substrings where found
            If (lngStartPosition > 0 And lngFirstPossibleEndPosition > 0 And lngSecondPossibleEndPosition > 0) Then
                blnFound = True
                'if the citations are located before the citation style in the block of text
                If lngStartPosition < lngFirstPossibleEndPosition Then
                    'gets the full information of all citations
                    strAllCitationsFullInformation = Mid(strWordOpenXML, lngStartPosition, lngFirstPossibleEndPosition - lngStartPosition)
                'if the citation style is located before the citations in the block of text
                Else
                    'gets the full information of all citations
                    strAllCitationsFullInformation = Mid(strWordOpenXML, lngStartPosition, lngSecondPossibleEndPosition - lngStartPosition)
                End If
                'replaces all '&quot;' by '"' to handle the string more easily
                strAllCitationsFullInformation = Replace(strAllCitationsFullInformation, "&quot;", Chr(34))
            End If
    End Select


    'if no information could be found
    If Not blnFound Then
        MsgBox "Could not find the full information of all citations." & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getAllCitationsFullInformation(intMendeleyVersion)"

        'stops the execution
        End
    End If


    'returns the string that contains all the information of all citations (when Mendeley Reference Manager 2.x is available)
    GAUG_getAllCitationsFullInformation = strAllCitationsFullInformation

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-23                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getCitationFullInfo(ByVal intMendeleyVersion As Integer, ByVal strAllCitationsFullInformation As String,                   **
'**     ByVal fldCitation As Field, ByVal ccCitation As ContentControl) As String                                                             **
'**                                                                                                                                           **
'**  Finds and returns the full information of a particular citation.                                                                         **
'**                                                                                                                                           **
'**  Parameter intMendeleyVersion can have two different values:                                                                              **
'**  1:                                                                                                                                       **
'**     Use version 1.x of Mendeley Desktop                                                                                                   **
'**  2:                                                                                                                                       **
'**     Use version 2.x of Mendeley Reference Manager                                                                                         **
'**  Parameter strAllCitationsFullInformation is a string that contains all information of all citations                                      **
'**     (when Mendeley Reference Manager 2.x is used)                                                                                         **
'**  Parameter fldCitation is the citation's field                                                                                            **
'**     (when Mendeley Desktop 1.x is used)                                                                                                   **
'**  Parameter ccCitation is the citation's content control                                                                                   **
'**     (when Mendeley Reference Manager 2.x is used)                                                                                         **
'**                                                                                                                                           **
'**  RETURNS: A string that contains the full information of the citation.                                                                    **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_getCitationFullInfo(ByVal intMendeleyVersion As Integer, ByVal strAllCitationsFullInformation As String, ByVal fldCitation As Field, ByVal ccCitation As ContentControl) As String

    Dim lngStartPositionOfCitation, lngEndPositionOfCitation, lngLengthOfCitation As Long
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
                'the full information of the citation is inside the citation's field
                strCitationFullInfo = fldCitation.Code
            End If

        'Mendeley Reference Manager 2.x is installed
        Case 2
            'if the citation's content control is not empty
            If Not (ccCitation Is Nothing) Then
                'initializes the start position (the start position, in the text, of the citation we need)
                lngStartPositionOfCitation = 1
                'iterates over all citations in the block of text
                Do
                    'gets the initial position of the text that contains the full information of the next citation
                    '(citations start with '{"citationID":"MENDELEY_CITATION_')
                    lngEndPositionOfCitation = InStr(lngStartPositionOfCitation, strAllCitationsFullInformation, "{" & Chr(34) & "citationID" & Chr(34) & ":" & Chr(34) & "MENDELEY_CITATION_")

                    'calculates the length of the current citation
                    If lngEndPositionOfCitation = 0 Then
                        lngLengthOfCitation = Len(strAllCitationsFullInformation)
                    Else
                        lngLengthOfCitation = lngEndPositionOfCitation - lngStartPositionOfCitation
                    End If

                    'if the tag of the searched citation is in the current match
                    If InStr(1, Mid(strAllCitationsFullInformation, lngStartPositionOfCitation, lngLengthOfCitation), ccCitation.Tag, vbTextCompare) > 0 Then
                        blnFound = True
                        'the full information of the citation is in this match
                        strCitationFullInfo = Mid(strAllCitationsFullInformation, lngStartPositionOfCitation, lngLengthOfCitation)
                        'exits loop
                        Exit Do
                    End If

                    'moves the start position for the next iteration
                    '(adding extra characters to prevent finding the same "next citation" next time)
                    lngStartPositionOfCitation = lngEndPositionOfCitation + 10

                Loop While lngEndPositionOfCitation > 0

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
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-04                                                                                                                **
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

    Dim lngLengthOfCitationItem As Long
    Dim varCitationItemsPositionsFromCitationFullInfo() As Variant
    Dim varCitationItemsFromCitationFullInfo() As Variant
    Dim i, intTotalCitationItems As Integer

    Dim objRegExpCitationItems As Object
    Dim colMatchesCitationItems As Object
    Dim objMatchVisibleCitationItemItem As Object


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

        Set objRegExpCitationItems = New GAUG_RegExp
        'sets case insensitivity
        objRegExpCitationItems.ignoreCase = False
        'sets global applicability
        objRegExpCitationItems.GlobalSearch = True

        'initializes the counter
        intTotalCitationItems = 0

        'builds the regular expression according to the version of Mendeley
        Select Case intMendeleyVersion
            'Mendeley Desktop 1.x is installed
            Case 1
                'sets the pattern to match '{"id":"ITEM'
                '(all individual citation items start with '{"id":"ITEM')
                objRegExpCitationItems.pattern = "\{\s*\" & Chr(34) & "id\" & Chr(34) & "\s*:\s*\" & Chr(34) & "ITEM"

            'Mendeley Reference Manager 2.x is installed
            Case 2
                'sets the pattern to match '{"id":"'
                '(all individual citation items start with '{"id":"'
                objRegExpCitationItems.pattern = "\{\s*\" & Chr(34) & "id\" & Chr(34) & "\s*:\s*\" & Chr(34)
        End Select


        'checks that the string can be compared
        If objRegExpCitationItems.Test(strCitationFullInfo) Then
            'gets the matches (individual citation items according to the regular expression)
            Set colMatchesCitationItems = objRegExpCitationItems.Execute(strCitationFullInfo)

            'treats all matches (all individual citation items)
            For Each objMatchVisibleCitationItemItem In colMatchesCitationItems
                'updates the counter to include this citation item
                intTotalCitationItems = intTotalCitationItems + 1
                'adds the start position of the full information of the citation item to the list
                ReDim Preserve varCitationItemsPositionsFromCitationFullInfo(1 To intTotalCitationItems)
                varCitationItemsPositionsFromCitationFullInfo(intTotalCitationItems) = objMatchVisibleCitationItemItem.FirstIndex
            Next objMatchVisibleCitationItemItem

            're-dimensions the array to store the full information of the citation items
            ReDim varCitationItemsFromCitationFullInfo(1 To intTotalCitationItems)

            'treats all individual citation items
            For i = 1 To intTotalCitationItems
                'if this is the last citation item
                If i = intTotalCitationItems Then
                    'the length is from the current citation item to the end of the string
                    lngLengthOfCitationItem = Len(strCitationFullInfo) - varCitationItemsPositionsFromCitationFullInfo(i)
                'if this is not the last citation item
                Else
                    'the length is from the current citation item to the next citation item
                    lngLengthOfCitationItem = varCitationItemsPositionsFromCitationFullInfo(i + 1) - varCitationItemsPositionsFromCitationFullInfo(i)
                End If 'if this is the last citation item

                'adds the full information of the citation item to the list
                varCitationItemsFromCitationFullInfo(i) = Mid(strCitationFullInfo, varCitationItemsPositionsFromCitationFullInfo(i) + 1, lngLengthOfCitationItem)
                'MsgBox Len(varCitationItemsFromCitationFullInfo(intTotalCitationItems))
            Next
        End If 'checks that the string can be compared

    End If 'if the citation's full info is not empty

    'returns the list of all items (individual citations within the field or content control) from the citation full information
    GAUG_getCitationItemsFromCitationFullInfo = varCitationItemsFromCitationFullInfo

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-08                                                                                                                **
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
Function GAUG_getAuthorsEditorsFromCitationItem(ByVal intMendeleyVersion As Integer, ByVal strAuthorEditor As String, ByVal strCitationItem As String) As Variant()

    Dim varAuthorsFromCitationItem() As Variant
    Dim intTotalAuthorsEditorsFromCitationItem As Integer
    Dim strFamilyName As String

    Dim objRegExpAuthorsFromCitationItem, objRegExpAuthorFamilyNamesFromCitationItem As Object
    Dim colMatchesAuthorsFromCitationItem, colMatchesAuthorFamilyNamesFromCitationItem As Object
    Dim objMatchAuthorFromCitationItem, objMatchAuthorFamilyNameFromCitationItem As Object


    'if the argument is not within valid versions
    If intMendeleyVersion < 1 Or intMendeleyVersion > 2 Then
        MsgBox "The version " & intMendeleyVersion & " of Mendeley's plugin is not valid." & vbCrLf & vbCrLf & _
            "Use version 1 or version 2 instead," & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getAuthorsEditorsFromCitationItem(intMendeleyVersion, strAuthorEditor, strCitationItem)"

        'stops the execution
        End
    End If

    'if the argument is not 'author' or 'editor'
    If Not (strAuthorEditor = "author" Or strAuthorEditor = "editor") Then
        MsgBox "The argument for strAuthorEditor is not valid." & vbCrLf & vbCrLf & _
            "Use " & Chr(34) & "author" & Chr(34) & " or " & Chr(34) & "editor" & Chr(34) & " instead," & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_getAuthorsEditorsFromCitationItem(intMendeleyVersion, strAuthorEditor, strCitationItem)"

        'stops the execution
        End
    End If


    'if the citation item's full info is not empty
    If Not (strCitationItem = "") Then

        Set objRegExpAuthorsFromCitationItem = New GAUG_RegExp
        'sets case insensitivity
        objRegExpAuthorsFromCitationItem.ignoreCase = False
        'sets global applicability
        objRegExpAuthorsFromCitationItem.GlobalSearch = True

        'builds the regular expression according to the version of Mendeley
        Select Case intMendeleyVersion
            'Mendeley Desktop 1.x is installed
            Case 1
                'sets the pattern to match everything from '"author":[' or '"editor":[' to (but not including) ']'
                'this gets the full list of authors from the citation item
                objRegExpAuthorsFromCitationItem.pattern = "\" & Chr(34) & strAuthorEditor & "\" & Chr(34) & "\s*:\s*\[.+?(?=\])"

            'Mendeley Reference Manager 2.x is installed
            Case 2
                'sets the pattern to match everything from '"author":[' or '"editor":[' to (but not including) ']'
                'this gets the full list of authors from the citation item
                objRegExpAuthorsFromCitationItem.pattern = "\" & Chr(34) & strAuthorEditor & "\" & Chr(34) & "\s*:\s*\[.+?(?=\])"
        End Select


        'checks that the string can be compared
        If objRegExpAuthorsFromCitationItem.Test(strCitationItem) Then
            'gets the matches (list of all authors as a single block of data)
            Set colMatchesAuthorsFromCitationItem = objRegExpAuthorsFromCitationItem.Execute(strCitationItem)

            'treats all matches (there should be at most one match, zero when editors are listed instead of the authors)
            For Each objMatchAuthorFromCitationItem In colMatchesAuthorsFromCitationItem

                Set objRegExpAuthorFamilyNamesFromCitationItem = New GAUG_RegExp
                'sets case insensitivity
                objRegExpAuthorFamilyNamesFromCitationItem.ignoreCase = False
                'sets global applicability
                objRegExpAuthorFamilyNamesFromCitationItem.GlobalSearch = True

                'initializes the counter
                intTotalAuthorsEditorsFromCitationItem = 0

                'builds the regular expression according to the version of Mendeley
                Select Case intMendeleyVersion
                    'Mendeley Desktop 1.x is installed
                    Case 1
                        'sets the pattern to match everything from '"family":"' to (but not including) '"'
                        'this gets the family names of authors from the citation item
                        objRegExpAuthorFamilyNamesFromCitationItem.pattern = "\" & Chr(34) & "family\" & Chr(34) & "\s*:\s*\" & Chr(34) & ".+?(?=\" & Chr(34) & ")"

                    'Mendeley Reference Manager 2.x is installed
                    Case 2
                        'sets the pattern to match everything from '{"family":"' to (but not including) '"'
                        'this gets the family names of authors from the citation item
                        objRegExpAuthorFamilyNamesFromCitationItem.pattern = "\{\s*\" & Chr(34) & "family\" & Chr(34) & "\s*:\s*\" & Chr(34) & ".+?(?=\" & Chr(34) & ")"
                End Select


                'checks that the string can be compared
                If objRegExpAuthorFamilyNamesFromCitationItem.Test(objMatchAuthorFromCitationItem.value) Then
                    'gets the matches (the family names of all authors)
                    Set colMatchesAuthorFamilyNamesFromCitationItem = objRegExpAuthorFamilyNamesFromCitationItem.Execute(objMatchAuthorFromCitationItem.value)

                    'treats all matches (the family name of the authors, if any)
                    For Each objMatchAuthorFamilyNameFromCitationItem In colMatchesAuthorFamilyNamesFromCitationItem
                        'gets only the family name, without the extra characters in the match
                        'from '{"family":"FamilyName' to just "FamilyName"
                        strFamilyName = Right(objMatchAuthorFamilyNameFromCitationItem.value, Len(objMatchAuthorFamilyNameFromCitationItem.value) - InStr(1, objMatchAuthorFamilyNameFromCitationItem.value, ":", vbTextCompare) - 1)
                        'updates the counter to include this family name of the author
                        intTotalAuthorsEditorsFromCitationItem = intTotalAuthorsEditorsFromCitationItem + 1
                        'adds the family name of the author to the list
                        ReDim Preserve varAuthorsFromCitationItem(1 To intTotalAuthorsEditorsFromCitationItem)
                        varAuthorsFromCitationItem(intTotalAuthorsEditorsFromCitationItem) = strFamilyName
                    Next objMatchAuthorFamilyNameFromCitationItem
                End If 'checks that the string can be compared


            Next objMatchAuthorFromCitationItem 'treats all matches (there should be at most one match, zero when editors are listed instead of the authors)

        End If 'checks that the string can be compared

    End If 'if the citation's full info is not empty


    'returns the list of the family names of the authors
    GAUG_getAuthorsEditorsFromCitationItem = varAuthorsFromCitationItem

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-04                                                                                                                **
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

    Dim objRegExpYearFromCitationItem As Object
    Dim colMatchesYearFromCitationItem As Object
    Dim objMatchYearFromCitationItem As Object


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

        Set objRegExpYearFromCitationItem = New GAUG_RegExp
        'sets case insensitivity
        objRegExpYearFromCitationItem.ignoreCase = False
        'sets global applicability
        objRegExpYearFromCitationItem.GlobalSearch = True

        'builds the regular expression according to the version of Mendeley
        Select Case intMendeleyVersion
            'Mendeley Desktop 1.x is installed
            Case 1
                'sets the pattern to match everything from '"issued":{"date-parts":[[' to (but not including) ']]' or ','
                'this gets only the year from the citation item, skips the month if present
                objRegExpYearFromCitationItem.pattern = "\" & Chr(34) & "issued\" & Chr(34) & "\s*:\s*\{\s*\" & Chr(34) & "date\-parts\" & Chr(34) & "\s*:\s*\[\s*\[.+?(?=((\s*\]\s*\])|(,)))"

            'Mendeley Reference Manager 2.x is installed
            Case 2
                'sets the pattern to match everything from '"issued":{"date-parts":[[' to (but not including) ']]' or ','
                'this gets only the year from the citation item, skips the month if present
                objRegExpYearFromCitationItem.pattern = "\" & Chr(34) & "issued\" & Chr(34) & "\s*:\s*\{\s*\" & Chr(34) & "date\-parts\" & Chr(34) & "\s*:\s*\[\s*\[.+?(?=((\s*\]\s*\])|(,)))"
        End Select


        'checks that the string can be compared
        If objRegExpYearFromCitationItem.Test(strCitationItem) Then
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
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-05                                                                                                                **
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


    'initializes the parts to empty string
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



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-17                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_getSafeStringForRegularExpressions(ByVal strOriginalString As String) As String                                            **
'**                                                                                                                                           **
'**  Replaces all instances of \.[]{}()<>*+-=!?^$| with the corresponding escaped character.                                                  **
'**                                                                                                                                           **
'**  Parameter strOriginalString is the original string                                                                                       **
'**                                                                                                                                           **
'**  RETURNS: The modified string after escaping the special characters.                                                                      **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_getSafeStringForRegularExpressions(ByVal strOriginalString As String) As String

    Dim arrSpecialCharacters(), strSpecialCharacter As Variant


    'defines the special characters, "\" must be first, otherwise it will be escaped multiple times
    arrSpecialCharacters = Array("\", ".", "[", "]", "{", "}", "(", ")", "<", ">", "*", "+", "-", "=", "!", "?", "^", "$", "|")

    'iterates over all characters in the original string
    For Each strSpecialCharacter In arrSpecialCharacters
        'escapes the current special character
        strOriginalString = Replace(strOriginalString, strSpecialCharacter, "\" & strSpecialCharacter)
    Next strSpecialCharacter


    'returns the modified original string
    GAUG_getSafeStringForRegularExpressions = strOriginalString

End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-09                                                                                                                **
'**                                                                                                                                           **
'**  Function GAUG_createHyperlinksForURLsInBibliography(ByVal intMendeleyVersion As Integer, ByVal fldBibliography As Field,                 **
'**     ByVal ccBibliography As ContentControl, arrNonDetectedURLs() As Variant) As Integer                                                   **
'**                                                                                                                                           **
'**  Generates the hyperlinks for the URLs in the bibliography inserted by Mendeley's plugin.                                                 **
'**                                                                                                                                           **
'**  Parameter intMendeleyVersion can have two different values:                                                                              **
'**  1:                                                                                                                                       **
'**     Use version 1.x of Mendeley Desktop                                                                                                   **
'**  2:                                                                                                                                       **
'**     Use version 2.x of Mendeley Reference Manager                                                                                         **
'**  Parameter fldBibliography is the bibliography's field                                                                                    **
'**     (when Mendeley Desktop 1.x is used)                                                                                                   **
'**  Parameter ccBibliography is the bibliography's content control                                                                           **
'**     (when Mendeley Reference Manager 2.x is used)                                                                                         **
'**  Parameter arrNonDetectedURLs specifies the URLs, not detected by the regular expression,                                                 **
'**     to generate the hyperlinks in the bibliography                                                                                        **
'**                                                                                                                                           **
'**  RETURNS: An integer with the number of hyperlinks that could not be created for the URLs.                                                **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Function GAUG_createHyperlinksForURLsInBibliography(ByVal intMendeleyVersion As Integer, ByVal fldBibliography As Field, ByVal ccBibliography As ContentControl, arrNonDetectedURLs() As Variant) As Integer

    Dim blnURLFound As Boolean
    Dim objRegExpURL As Object
    Dim colMatchesURL As Object
    Dim objMatchURL As Object
    Dim strURL, strSubStringOfURL As String
    Dim varNonDetectedURL As Variant
    Dim objCurrentFieldOrContentControl As Object
    Dim intTotalURLsWithoutHyperlink As Integer


    'if the argument is not within valid versions
    If intMendeleyVersion < 1 Or intMendeleyVersion > 2 Then
        MsgBox "The version " & intMendeleyVersion & " of Mendeley's plugin is not valid." & vbCrLf & vbCrLf & _
            "Use version 1 or version 2 instead," & vbCrLf & vbCrLf & _
            "Cannot continue creating hyperlinks.", _
            vbCritical, "GAUG_createHyperlinksForURLsInBibliography(intMendeleyVersion, fldBibliography, ccBibliography, arrNonDetectedURLs)"

        'stops the execution
        End
    End If


    'creates the object for regular expressions (to get all URLs in bibliography)
    Set objRegExpURL = New GAUG_RegExp
    'sets the pattern to match every URL in the bibliography (http, https or ftp)
    objRegExpURL.pattern = "((https?)|(ftp)):\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z0-9]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=\[\]\(\)<>;]*)"
    'sets case insensitivity
    objRegExpURL.ignoreCase = False
    'sets global applicability
    objRegExpURL.GlobalSearch = True


    Select Case intMendeleyVersion
        'Mendeley Desktop 1.x is installed
        Case 1
            'if the bibliography's field is not empty
            If Not (fldBibliography Is Nothing) Then
                'gets the field of the bibliography
                Set objCurrentFieldOrContentControl = fldBibliography
            Else
                'nothing to do
                GAUG_createHyperlinksForURLsInBibliography = 0
                Exit Function
            End If
        'Mendeley Reference Manager 2.x is installed
        Case 2
            'if the bibliography's content control is not empty
            If Not (ccBibliography Is Nothing) Then
                'gets the content control of the bibliography
                Set objCurrentFieldOrContentControl = ccBibliography
            Else
                'nothing to do
                GAUG_createHyperlinksForURLsInBibliography = 0
                Exit Function
            End If
    End Select


    'if the array of non detected URLs is not empty
    If Not Not arrNonDetectedURLs Then
        'generates the hyperlinks from the list of non detected URLs
        'the non detected URLs shall be done first or some conflicts may arise
        For Each varNonDetectedURL In arrNonDetectedURLs
            'prevents errors if the match is longer than 256 characters
            strSubStringOfURL = Left(CStr(varNonDetectedURL), 256)

            'according to the version of Mendeley
            Select Case intMendeleyVersion
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
                    .Text = strSubStringOfURL
                    .Execute
                    blnURLFound = .found
                End With

                If blnURLFound Then
                    'moves the selection, if necessary, to include the full match
                    Selection.MoveEnd Unit:=wdCharacter, Count:=Len(CStr(varNonDetectedURL)) - Len(strSubStringOfURL)

                    'checks that the full match is found
                    If Selection.Text = CStr(varNonDetectedURL) Then
                        blnURLFound = True
                    Else
                        'there is no more searching, the hyperlink for this URL will not be created
                        blnURLFound = False
                        intTotalURLsWithoutHyperlink = intTotalURLsWithoutHyperlink + 1
                    End If
                End If

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
        Next 'generates the hyperlinks from the list of non detected URLs
    End If 'if the array of non detected URLs is not empty

    'according to the version of Mendeley
    Select Case intMendeleyVersion
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
    If objRegExpURL.Test(Selection) Then
        'gets the matches (all URLs in the bibliography according to the regular expression)
        Set colMatchesURL = objRegExpURL.Execute(Selection)

        'treats all matches (all URLs in bibliography) to generate hyperlinks
        For Each objMatchURL In colMatchesURL

            'removes the last character if a period (".")
            If Right(objMatchURL.value, 1) = "." Then
                strURL = Left(objMatchURL.value, Len(objMatchURL.value) - 1)
            Else
                strURL = objMatchURL.value
            End If

            'prevents errors if the match is longer than 256 characters
            strSubStringOfURL = Left(strURL, 256)

            'according to the version of Mendeley
            Select Case intMendeleyVersion
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
                    .Text = strSubStringOfURL
                    .Execute
                    blnURLFound = .found
                End With

                If blnURLFound Then
                    'moves the selection, if necessary, to include the full match
                    Selection.MoveEnd Unit:=wdCharacter, Count:=Len(strURL) - Len(strSubStringOfURL)

                    'checks that the full match is found
                    If Selection.Text = strURL Then
                        blnURLFound = True
                    Else
                        'there is no more searching, the hyperlink for this URL will not be created
                        blnURLFound = False
                        intTotalURLsWithoutHyperlink = intTotalURLsWithoutHyperlink + 1
                    End If
                End If

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

        Next 'treats all matches (all URLs in bibliography) to generate hyperlinks
    End If 'checks that the string can be compared


    'returns the number of URLs for which the hyperlinks could not be created
    GAUG_createHyperlinksForURLsInBibliography = intTotalURLsWithoutHyperlink
End Function



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-30                                                                                                                **
'**                                                                                                                                           **
'**  Sub GAUG_createHyperlinksForCitationsAPA()                                                                                               **
'**                                                                                                                                           **
'**  Generates the bookmarks in the bibliography inserted by Mendeley's plugin.                                                               **
'**  Links the citations inserted by Mendeley's plugin to the corresponding entry in the bibliography inserted by Mendeley's plugin.          **
'**  Generates the hyperlinks for the URLs in the bibliography inserted by Mendeley's plugin.                                                 **
'**  Only for APA CSL citation style.                                                                                                         **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Sub GAUG_createHyperlinksForCitationsAPA()

    Dim intAvailableMendeleyVersion As Integer, intUseMendeleyVersion As Integer
    Dim documentSection As Section
    Dim blnFound As Boolean, blnBibliographyFound As Boolean, blnCitationFound As Boolean, blnReferenceEntryFound As Boolean, blnCitationEntryFound As Boolean, blnGenerateHyperlinksForURLs As Boolean, blnURLFound As Boolean
    Dim intReferenceNumber As Integer
    Dim objRegExpBibliographyEntries As Object, objRegExpVisibleCitationItems As Object, objRegExpFindBibliographyEntry As Object, objRegExpFindVisibleCitationItem As Object, objRegExpURL As Object
    Dim colMatchesBibliographyEntries As Object, colMatchesVisibleCitationItems As Object, colMatchesFindBibliographyEntry As Object, colMatchesFindVisibleCitationItem As Object, colMatchesURL As Object
    Dim objMatchBibliographyEntry As Object, objMatchVisibleCitationItem As Object, objMatchFindBibliographyEntry As Object, objMatchURL As Object
    Dim strBookmarkInBibliography As Variant, arrStrBookmarksInBibliography() As String
    Dim strTempMatch As String, strSubStringOfTempMatch As String, strLastAuthorsOrEditors As String
    Dim strTypeOfExecution As String
    Dim blnMabEntwickeltSich As Boolean
    Dim stlStyleInDocument As Word.Style
    Dim strStyleForTitleOfBibliography As String
    Dim blnStyleForTitleOfBibliographyFound As Boolean
    Dim strURL As String, strSubStringOfURL As String
    Dim arrNonDetectedURLs() As Variant, varNonDetectedURL As Variant
    Dim strDoHyperlinksExist As String
    Dim objCurrentFieldOrContentControl As Object
    Dim strAllCitationsFullInformation As String, strCitationFullInfo As String
    Dim varCitationItemsFromCitationFullInfo() As Variant
    Dim varPartsFromVisibleCitationItem() As Variant
    Dim varAuthorsFromCitationItem() As Variant
    Dim varEditorsFromCitationItem() As Variant
    Dim varYearFromCitationItem As String
    Dim intAuthorFromCitationItem As Integer, intEditorFromCitationItem As Integer
    Dim intCitationItemFromCitationFullInfo As Integer
    Dim strOrphanCitationItems As String
    Dim varFieldsOrContentControls As Variant
    Dim currentPosition As range
    Dim intTotalURLsWithoutHyperlink As Integer
    Dim strBibliographyFullEntries() As String
    Dim lngCurrentBibliographyFullEntry As Long
    Dim strVisibleTextOfCurrentCitation As String


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
    strStyleForTitleOfBibliography = "Titre de dernière section"

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
    'creates the object for regular expressions (to get the entry's text in the bibliography to create the bookmark)
    Set objRegExpBibliographyEntries = New GAUG_RegExp
    'sets the pattern to match the reference entry
    '(all text from the beginning of the string until a year between parentheses is found)
    objRegExpBibliographyEntries.pattern = "^.*\.\s\(\d\d\d\d[a-zA-Z]?\)"
    'sets case insensitivity
    objRegExpBibliographyEntries.ignoreCase = False
    'sets global applicability
    objRegExpBibliographyEntries.GlobalSearch = True
    'creates the object for regular expressions (to get all URLs in bibliography)
    Set objRegExpURL = New GAUG_RegExp
    'sets the pattern to match every URL in the bibliography (http, https or ftp)
    objRegExpURL.pattern = "((https?)|(ftp)):\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z0-9]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=\[\]\(\)<>;]*)"
    'sets case insensitivity
    objRegExpURL.ignoreCase = False
    'sets global applicability
    objRegExpURL.GlobalSearch = True

    'initializes the flag
    blnBibliographyFound = False
    'initializes the counter for URLs without hyperlink generated for them
    intTotalURLsWithoutHyperlink = 0


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
                        'checks if it is the bibliography
                        If objCurrentFieldOrContentControl.Type = wdContentControlRichText And Trim(objCurrentFieldOrContentControl.Tag) = "MENDELEY_BIBLIOGRAPHY" Then
                            blnBibliographyFound = True
                        End If
                End Select

                'if it is the bibliography
                If blnBibliographyFound Then
                    'start the numbering
                    intReferenceNumber = 1

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

                    'separates the bibliography in individual full entries
                    strBibliographyFullEntries = Split(Selection, Chr(13))

                    'creates an array for the matches in the bibliography that will have a bookmark
                    ReDim arrStrBookmarksInBibliography(1 To UBound(strBibliographyFullEntries) - LBound(strBibliographyFullEntries) + 1)


                    'treats all matches (all full entries in bibliography) to generate bookmarks
                    For lngCurrentBibliographyFullEntry = LBound(strBibliographyFullEntries) To UBound(strBibliographyFullEntries)
                        'checks that the string can be compared
                        If objRegExpBibliographyEntries.Test(strBibliographyFullEntries(lngCurrentBibliographyFullEntry)) Then
                            'gets the matches (the single entry according to the regular expression)
                            Set colMatchesBibliographyEntries = objRegExpBibliographyEntries.Execute(strBibliographyFullEntries(lngCurrentBibliographyFullEntry))

                            'if a match was found in this bibliography full entry (there should always be only one)
                            If colMatchesBibliographyEntries.Count = 1 Then
                                'treats all matches (should be only one) to generate bookmarks
                                '(we have to find AGAIN every entry to select it and create the bookmark)
                                For Each objMatchBibliographyEntry In colMatchesBibliographyEntries
                                    'removes the carriage return from match, if necessary
                                    strTempMatch = Replace(objMatchBibliographyEntry.value, Chr(13), "")

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
                                            name:="GAUG_SignetBibliographie_" & format(CStr(intReferenceNumber), "00#"), _
                                            range:=Selection.range
                                        'adds the entry to the lists of bookmarks, to be used later when linking the citations to the bibliography
                                        arrStrBookmarksInBibliography(intReferenceNumber) = Selection.range.Text
                                    End If

                                    'continues with the next number
                                    intReferenceNumber = intReferenceNumber + 1

                                Next 'treats all matches (all entries in bibliography) to generate bookmarks
                            End If 'if a match was found in this bibliography full entry (there should always be only one)
                        End If 'checks that the string can be compared
                    Next 'treats all matches (all full entries in bibliography) to generate bookmarks


                    'by now, we have created all bookmarks and have all entries in arrStrBookmarksInBibliography
                    'for future use when creating the hyperlinks

                    'generates the hyperlinks for the URLs in the bibliography, if required
                    If blnGenerateHyperlinksForURLs Then

                        'according to the version of Mendeley
                        Select Case intAvailableMendeleyVersion
                            'Mendeley Desktop 1.x is installed
                            Case 1
                                'creates the hyperlinks for the URLs in the bibliography
                                intTotalURLsWithoutHyperlink = GAUG_createHyperlinksForURLsInBibliography(intAvailableMendeleyVersion, objCurrentFieldOrContentControl, Nothing, arrNonDetectedURLs)
                            'Mendeley Reference Manager 2.x is installed
                            Case 2
                                'creates the hyperlinks for the URLs in the bibliography
                                intTotalURLsWithoutHyperlink = GAUG_createHyperlinksForURLsInBibliography(intAvailableMendeleyVersion, Nothing, objCurrentFieldOrContentControl, arrNonDetectedURLs)
                        End Select
                    End If 'hyperlinks for URLs in bibliography

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
    'creates the objects for regular expressions (to get all entries in current citation and position of citation entry in bibliography)
    Set objRegExpVisibleCitationItems = New GAUG_RegExp
    Set objRegExpFindBibliographyEntry = New GAUG_RegExp
    Set objRegExpFindVisibleCitationItem = New GAUG_RegExp
    'sets the pattern to match every citation entry (with or without authors) in the visible text of the current field or content control
    '(the year of publication is always present, authors may not be)
    '(all text non starting by "(" or "," or ";" followed by non digits until a year is found)
    objRegExpVisibleCitationItems.pattern = "[^(\(|,|;)][^0-9]*\d\d\d\d[a-zA-Z]?"
    'sets case insensitivity
    objRegExpVisibleCitationItems.ignoreCase = False
    objRegExpFindBibliographyEntry.ignoreCase = False
    objRegExpFindVisibleCitationItem.ignoreCase = False
    'sets global applicability
    objRegExpVisibleCitationItems.GlobalSearch = True
    objRegExpFindBibliographyEntry.GlobalSearch = True
    objRegExpFindVisibleCitationItem.GlobalSearch = True

    strOrphanCitationItems = ""


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

                'gets the visible text of the current citation, excluding the parenthesis (used to make sure all citation items are treated)
                strVisibleTextOfCurrentCitation = Selection
                If Left(strVisibleTextOfCurrentCitation, 1) = "(" Then
                    strVisibleTextOfCurrentCitation = Right(strVisibleTextOfCurrentCitation, Len(strVisibleTextOfCurrentCitation) - 1)
                End If
                If Right(strVisibleTextOfCurrentCitation, 1) = ")" Then
                    strVisibleTextOfCurrentCitation = Left(strVisibleTextOfCurrentCitation, Len(strVisibleTextOfCurrentCitation) - 1)
                Else
                    If Left(Right(strVisibleTextOfCurrentCitation, 2), 1) = ")" Then
                        strVisibleTextOfCurrentCitation = Left(strVisibleTextOfCurrentCitation, Len(strVisibleTextOfCurrentCitation) - 2)
                    End If
                End If

                'checks that the string can be compared
                If objRegExpVisibleCitationItems.Test(Selection) Then
                    'gets the matches (all entries in the citation according to the regular expression)
                    Set colMatchesVisibleCitationItems = objRegExpVisibleCitationItems.Execute(Selection)

                    'treats all matches (all entries in citation) to generate hyperlinks
                    For Each objMatchVisibleCitationItem In colMatchesVisibleCitationItems
                        'gets the list of all parts (authors or editors, year, and letter of year if present) from the visible citation item (entry in visible text of citation)
                        'position 1 is the authors or editors (when present)
                        'position 2 is the issue year
                        'position 3 is the letter after the issue year (when present)
                        varPartsFromVisibleCitationItem = GAUG_getPartsFromVisibleCitationItem(objMatchVisibleCitationItem.value)

                        'when citations are merged, they are ordered by the authors' family names
                        'the position of the citation in the visible text may not correspond to the position in the citation hidden data,
                        'we need to find the entry, but we may not have the authors's family names :(

                        'if the current match has authors' family names (not only the year) (it could happen that they are the editors' family names, but we do not know yet)
                        'we keep them stored for future use if next citation item DOES NOT include them
                        If Len(varPartsFromVisibleCitationItem(1)) > 0 Then
                            strLastAuthorsOrEditors = varPartsFromVisibleCitationItem(1)
                        End If

                        'checks if the list of all items from the citation full information is not empty (for info on 'Not Not' see https://riptutorial.com/excel-vba/example/30824/check-if-array-is-initialized--if-it-contains-elements-or-not--)
                        If Not Not varCitationItemsFromCitationFullInfo Then
                            'treats all citation items from the citation full information (to find which one corresponds to the current visible citation item being treated)
                            For intCitationItemFromCitationFullInfo = 1 To UBound(varCitationItemsFromCitationFullInfo)
                                'gets the list of authors (if available) from the citation item
                                varAuthorsFromCitationItem = GAUG_getAuthorsEditorsFromCitationItem(intAvailableMendeleyVersion, "author", varCitationItemsFromCitationFullInfo(intCitationItemFromCitationFullInfo))
                                'gets the list of editors (if available) from the citation item
                                varEditorsFromCitationItem = GAUG_getAuthorsEditorsFromCitationItem(intAvailableMendeleyVersion, "editor", varCitationItemsFromCitationFullInfo(intCitationItemFromCitationFullInfo))
                                'gets the year of issue from the citation item
                                varYearFromCitationItem = GAUG_getYearFromCitationItem(intAvailableMendeleyVersion, varCitationItemsFromCitationFullInfo(intCitationItemFromCitationFullInfo))

                                'initializes the regular expressions
                                objRegExpFindBibliographyEntry.pattern = ""
                                objRegExpFindVisibleCitationItem.pattern = ""

                                'if the citation item has authors (instead of editors)
                                If Not Not varAuthorsFromCitationItem Then
                                    For intAuthorFromCitationItem = 1 To UBound(varAuthorsFromCitationItem)
                                        'gets the last name of the author and adds it to the regular expression (used to find the entry in the bibliography)
                                        objRegExpFindBibliographyEntry.pattern = objRegExpFindBibliographyEntry.pattern & GAUG_getSafeStringForRegularExpressions(varAuthorsFromCitationItem(intAuthorFromCitationItem)) & ".*" '",?[A-Z\.\s]*[&,\s]*"
                                        'creates another regular expression to match the entry in the hidden data with the citation item in the visible text
                                        If intAuthorFromCitationItem = 1 Then
                                            'if this is the first author, there could be only two and "&" is used, otherwise "," is used
                                            objRegExpFindVisibleCitationItem.pattern = objRegExpFindVisibleCitationItem.pattern & GAUG_getSafeStringForRegularExpressions(varAuthorsFromCitationItem(intAuthorFromCitationItem)) & "\s?(&|,)?\s?"
                                        Else
                                            'if this is not the first author, "," is used to separate them
                                            'if this is not the first author of many, this could have been replaced by "et al." in the current citation item
                                            'includes the part to check for "et al." (only for the visible citation item, the entry in the bibliography has the full list)
                                            objRegExpFindVisibleCitationItem.pattern = objRegExpFindVisibleCitationItem.pattern & "((et al\..*)|(" & GAUG_getSafeStringForRegularExpressions(varAuthorsFromCitationItem(intAuthorFromCitationItem)) & ",?\s?"
                                        End If
                                    Next
                                    'closes the parenthesis in the pattern if more than one author
                                    If UBound(varAuthorsFromCitationItem) > 1 Then
                                        'adds "))" for each of the authors, excluding the first one
                                        objRegExpFindVisibleCitationItem.pattern = objRegExpFindVisibleCitationItem.pattern & Replace(Space(UBound(varAuthorsFromCitationItem) - 1), " ", "))")
                                    End If

                                'but if no authors were found (like with a book with only editors), we use editors instead
                                Else
                                    'if the citation item has editors
                                    If Not Not varEditorsFromCitationItem Then
                                        For intEditorFromCitationItem = 1 To UBound(varEditorsFromCitationItem)
                                            'gets the last name of the editor and adds it to the regular expression (used to find the entry in the bibliography)
                                            objRegExpFindBibliographyEntry.pattern = objRegExpFindBibliographyEntry.pattern & GAUG_getSafeStringForRegularExpressions(varEditorsFromCitationItem(intEditorFromCitationItem)) & ".*" '",?[A-Z\.\s]*[&,\s]*"
                                            'creates another regular expression to match the entry in the hidden data with the citation item in the visible text
                                            If intEditorFromCitationItem = 1 Then
                                                'if this is the first editor, there could be only two and "&" is used, otherwise "," is used
                                                objRegExpFindVisibleCitationItem.pattern = objRegExpFindVisibleCitationItem.pattern & GAUG_getSafeStringForRegularExpressions(varEditorsFromCitationItem(intEditorFromCitationItem)) & "\s?(&|,)?\s?"
                                            Else
                                                'if this is not the first editor, "," is used to separate them
                                                'if this is not the first editor of many, this could have been replaced by "et al." in the current citation item
                                                'includes the part to check for "et al." (only for the visible citation item, the entry in the bibliography has the full list)
                                                objRegExpFindVisibleCitationItem.pattern = objRegExpFindVisibleCitationItem.pattern & "((et al\..*)|(" & GAUG_getSafeStringForRegularExpressions(varEditorsFromCitationItem(intEditorFromCitationItem)) & ",?\s?"
                                            End If
                                        Next
                                        'because the citation has editors, we make sure the text '(Eds.)' or '(Ed.)' is present in the bibliography
                                        objRegExpFindBibliographyEntry.pattern = objRegExpFindBibliographyEntry.pattern & "\(Eds?\.\)\.\s*"
                                        'closes the parenthesis in the pattern if more than one editor
                                        If UBound(varEditorsFromCitationItem) > 1 Then
                                            'adds "))" for each of the editors, excluding the first one
                                            objRegExpFindVisibleCitationItem.pattern = objRegExpFindVisibleCitationItem.pattern & Replace(Space(UBound(varEditorsFromCitationItem) - 1), " ", "))")
                                        End If
                                    End If
                                End If

                                'if authors or editors exist
                                If (objRegExpFindBibliographyEntry.pattern <> "" And objRegExpFindVisibleCitationItem.pattern <> "") Then

                                    'finishes the patterns including the year and the letter shown in the visible citation item
                                    objRegExpFindVisibleCitationItem.pattern = objRegExpFindVisibleCitationItem.pattern & varYearFromCitationItem & varPartsFromVisibleCitationItem(3)
                                    objRegExpFindBibliographyEntry.pattern = objRegExpFindBibliographyEntry.pattern & "\(" & varYearFromCitationItem & varPartsFromVisibleCitationItem(3) & "\)"
                                    'MsgBox objMatchVisibleCitationItem.value & " -> *" & _
                                    '    varPartsFromVisibleCitationItem(1) & "*" & varPartsFromVisibleCitationItem(2) & "*" & varPartsFromVisibleCitationItem(3) & "*" & vbCrLf & vbCrLf & _
                                    '    objRegExpFindVisibleCitationItem.pattern & vbCrLf & vbCrLf & _
                                    '    objRegExpFindBibliographyEntry.pattern
                                Else
                                    objRegExpFindVisibleCitationItem.pattern = "Error: No authors or editors"
                                    objRegExpFindBibliographyEntry.pattern = "Error: No authors or editors"
                                End If

                                'if the current visible citation item has authors's family names (not only the year)
                                If Len(varPartsFromVisibleCitationItem(1)) > 0 Then
                                    'checks if this item from the citation full information corresponds to the visible citation item being treated
                                    Set colMatchesFindVisibleCitationItem = objRegExpFindVisibleCitationItem.Execute(objMatchVisibleCitationItem.value)
                                Else
                                    'checks if this item from the citation full information corresponds to the visible citation item being treated
                                    Set colMatchesFindVisibleCitationItem = objRegExpFindVisibleCitationItem.Execute(strLastAuthorsOrEditors & ", " & objMatchVisibleCitationItem.value)
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
                        'it also checks the case when both authors and editors are not present
                        If colMatchesFindVisibleCitationItem.Count = 0 Then
                            'cleans the regular expression as no matches were found
                            objRegExpFindBibliographyEntry.pattern = "Error: Citation not found"
                        End If

                        'at this point, the regular expression to find the entry in the bibliography is ready
                        'it is time to find the citation entry in the bibliography and link the current visible citation item

                        'MsgBox "The visible citation item" & vbCrLf & _
                        '    objMatchVisibleCitationItem.value & vbCrLf & _
                        '    "matches:" & vbCrLf & vbCrLf & _
                        '    objRegExpFindVisibleCitationItem.pattern & vbCrLf & _
                        '    objRegExpFindBibliographyEntry.pattern

                        'initializes the position
                        intReferenceNumber = 1
                        'finds the position of the citation entry in the list of references in the bibliography
                        blnReferenceEntryFound = False
                        For Each strBookmarkInBibliography In arrStrBookmarksInBibliography
                            'MsgBox ("Searching for citation in bibliography:" & vbCrLf & vbCrLf & "Using..." & vbCrLf & objRegExpFindBibliographyEntry.pattern & vbCrLf & strBookmarkInBibliography)
                            'gets the matches, if any, to check if this reference entry corresponds to the visible citation item being treated
                            Set colMatchesFindBibliographyEntry = objRegExpFindBibliographyEntry.Execute(CStr(strBookmarkInBibliography))
                            'if this is the corresponding reference entry
                            'Verify for MabEntwickeltSich: perhaps a more strict verification is needed
                            If colMatchesFindBibliographyEntry.Count > 0 Then
                                blnReferenceEntryFound = True
                                Exit For
                            End If
                            'continues with the next number
                            intReferenceNumber = intReferenceNumber + 1
                        Next

                        'at this point we also have the position (intReferenceNumber) in the bibliography, we are ready to create the hyperlink
                        'the position is issued to link to the bookmark 'GAUG_SignetBibliographie_<position>'

                        'if reference entry was found (shall always find it), creates the hyperlink
                        If blnReferenceEntryFound Then
                            'MsgBox ("Citation was found in the bibliography" & vbCrLf & vbCrLf & colMatchesFindBibliographyEntry.Item(0).value)
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
                                'better to use normal hyperlink:
                                Selection.Hyperlinks.Add Anchor:=Selection.range, _
                                    Address:="", SubAddress:="GAUG_SignetBibliographie_" & format(CStr(intReferenceNumber), "00#"), _
                                    ScreenTip:=""
                            End If
                        Else
                            'if the visible citation item could not be linked to an entry in the bibliography
                            strOrphanCitationItems = strOrphanCitationItems & Trim(objMatchVisibleCitationItem.value) & vbCrLf
                        End If

                        'at this point current citation entry is linked to corresponding reference in bibliography

                        'if the current match has authors' family names (not only the year)
                        If Len(varPartsFromVisibleCitationItem(1)) > 0 Then
                            'removes the current citation item from the visible text variable because it is already treated (if it is NOT the first one)
                            strVisibleTextOfCurrentCitation = Replace(strVisibleTextOfCurrentCitation, "; " & varPartsFromVisibleCitationItem(1) & ", " & varPartsFromVisibleCitationItem(2) & varPartsFromVisibleCitationItem(3), "", 1, 1)
                            'removes the current citation item from the visible text variable because it is already treated (if it is the first one)
                            strVisibleTextOfCurrentCitation = Replace(strVisibleTextOfCurrentCitation, varPartsFromVisibleCitationItem(1) & ", " & varPartsFromVisibleCitationItem(2) & varPartsFromVisibleCitationItem(3), "", 1, 1)
                        Else
                            'removes the current citation item from the visible text variable because it is already treated
                            strVisibleTextOfCurrentCitation = Replace(strVisibleTextOfCurrentCitation, ", " & varPartsFromVisibleCitationItem(2) & varPartsFromVisibleCitationItem(3), "", 1, 1)
                        End If

                    Next 'treats all matches (all entries in citation) to generate hyperlinks

                End If 'checks that the string can be compared

                'if the visible text variable is not empty, some citation item was not linked
                If Len(Trim(strVisibleTextOfCurrentCitation)) > 0 Then
                    'MsgBox "Visible text of citation: " & strVisibleTextOfCurrentCitation
                    strOrphanCitationItems = strOrphanCitationItems & strVisibleTextOfCurrentCitation & vbCrLf
                End If

            End If 'if it is a citation
        Next 'checks all fields or content controls
    Next 'documentSection

    'at this point all citations are linked to their corresponding reference in bibliography

    'if orphan citations exist
    If Len(strOrphanCitationItems) > 0 Then
        MsgBox "Orphan citation entries found:" & vbCrLf & vbCrLf & _
            strOrphanCitationItems & vbCrLf & _
            "Remove them from the document or manually create the bookmarks and hyperlinks!", _
            vbExclamation, "GAUG_createHyperlinksForCitationsAPA()"
    End If

    'if hyperlinks where not generated for the URLs in the bibliography
    If intTotalURLsWithoutHyperlink > 0 Then
        MsgBox "The bibliography contains " & intTotalURLsWithoutHyperlink & " URLs for which" & vbCrLf & _
            "the macros could not create the hyperlinks." & vbCrLf & vbCrLf & _
            "Create the hyperlinks manually by selecting every URL," & vbCrLf & _
            "then go to the menu Insert and click on Link.", _
            vbExclamation, "GAUG_createHyperlinksForCitationsAPA()"
        End
    End If

    'returns to original position in the document
    currentPosition.Select

    'reenables the screen updating
    Application.ScreenUpdating = True

End Sub



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-09                                                                                                                **
'**                                                                                                                                           **
'**  Sub GAUG_createHyperlinksForCitationsIEEE()                                                                                              **
'**                                                                                                                                           **
'**  Generates the bookmarks in the bibliography inserted by Mendeley's plugin.                                                               **
'**  Links the citations inserted by Mendeley's plugin to the corresponding entry in the bibliography inserted by Mendeley's plugin.          **
'**  Generates the hyperlinks for the URLs in the bibliography inserted by Mendeley's plugin.                                                 **
'**  Only for IEEE CSL citation style.                                                                                                        **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Sub GAUG_createHyperlinksForCitationsIEEE()

    Dim intAvailableMendeleyVersion, intUseMendeleyVersion As Integer
    Dim documentSection As Section
    Dim sectionField As Field
    Dim blnFound, blnCitationFound, blnBibliographyFound, blnReferenceNumberFound, blnCitationNumberFound, blnGenerateHyperlinksForURLs, blnURLFound As Boolean
    Dim intReferenceNumber, intCitationNumber As Integer
    Dim objRegExpVisibleCitationItems, objRegExpURL As Object
    Dim colMatchesVisibleCitationItems, colMatchesURL As Object
    Dim objMatchVisibleCitationItem, objMatchURL As Object
    Dim blnIncludeSquareBracketsInHyperlinks As Boolean
    Dim strTypeOfExecution As String
    Dim blnMabEntwickeltSich As Boolean
    Dim stlStyleInDocument As Word.Style
    Dim strStyleForTitleOfBibliography As String
    Dim blnStyleForTitleOfBibliographyFound As Boolean
    Dim strURL, strSubStringOfURL As String
    Dim arrNonDetectedURLs(), varNonDetectedURL As Variant
    Dim strDoHyperlinksExist As String
    Dim objCurrentFieldOrContentControl As Object
    Dim strOrphanCitationItems As String
    Dim varFieldsOrContentControls As Variant
    Dim currentPosition As range
    Dim blnIsSingleCitation As Boolean
    Dim intTotalURLsWithoutHyperlink As Integer


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
    strStyleForTitleOfBibliography = "Titre de dernière section"

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
    'if set to True, then the square brackets will be part of the hyperlink
    blnIncludeSquareBracketsInHyperlinks = False

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
            vbCritical, "GAUG_createHyperlinksForCitationsIEEE()"

        'stops the execution
        End
    End If

    'gets the version of Mendeley (autodetect or specified)
    intAvailableMendeleyVersion = GAUG_getAvailableMendeleyVersion(intUseMendeleyVersion)

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
    'creates the object for regular expressions (to get all URLs in bibliography)
    Set objRegExpURL = New GAUG_RegExp
    'sets the pattern to match every URL in the bibliography (http, https or ftp)
    objRegExpURL.pattern = "((https?)|(ftp)):\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z0-9]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=\[\]\(\)<>;]*)"
    'sets case insensitivity
    objRegExpURL.ignoreCase = False
    'sets global applicability
    objRegExpURL.GlobalSearch = True

    'initializes the flag
    blnBibliographyFound = False
    'initializes the counter for URLs without hyperlink generated for them
    intTotalURLsWithoutHyperlink = 0


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
                        'checks if it is the bibliography
                        If objCurrentFieldOrContentControl.Type = wdContentControlRichText And Trim(objCurrentFieldOrContentControl.Tag) = "MENDELEY_BIBLIOGRAPHY" Then
                            blnBibliographyFound = True
                        End If
                End Select

                'if it is the bibliography
                If blnBibliographyFound Then
                    'start the numbering
                    intReferenceNumber = 1
                    Do
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

                        'finds and selects the text of the number of the reference
                        With Selection.Find
                            .Forward = True
                            .Wrap = wdFindStop
                            .Text = "[" & CStr(intReferenceNumber) & "]"
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
                                    .Text = CStr(intReferenceNumber)
                                    .Execute
                                    blnReferenceNumberFound = .found
                                End With
                            End If

                            'creates the bookmark
                            Selection.Bookmarks.Add _
                                name:="GAUG_SignetBibliographie_" & format(CStr(intReferenceNumber), "00#"), _
                                range:=Selection.range
                        End If

                        'continues with the next number
                        intReferenceNumber = intReferenceNumber + 1

                    'while numbers of references are found
                    Loop While (blnReferenceNumberFound)

                    'by now, we have created all bookmarks with sequential numbers
                    'for future use when creating the hyperlinks

                    'generates the hyperlinks for the URLs in the bibliography, if required
                    If blnGenerateHyperlinksForURLs Then

                        'according to the version of Mendeley
                        Select Case intAvailableMendeleyVersion
                            'Mendeley Desktop 1.x is installed
                            Case 1
                                'creates the hyperlinks for the URLs in the bibliography
                                intTotalURLsWithoutHyperlink = GAUG_createHyperlinksForURLsInBibliography(intAvailableMendeleyVersion, objCurrentFieldOrContentControl, Nothing, arrNonDetectedURLs)
                            'Mendeley Reference Manager 2.x is installed
                            Case 2
                                'creates the hyperlinks for the URLs in the bibliography
                                intTotalURLsWithoutHyperlink = GAUG_createHyperlinksForURLsInBibliography(intAvailableMendeleyVersion, Nothing, objCurrentFieldOrContentControl, arrNonDetectedURLs)
                        End Select
                    End If 'hyperlinks for URLs in bibliography

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
    'creates the object for regular expressions (to get all entries in current citation)
    Set objRegExpVisibleCitationItems = New GAUG_RegExp
    'sets the pattern to match every citation entry in the visible text of the current field or content control
    'it should be "[" + Number + "]"
    objRegExpVisibleCitationItems.pattern = "\[[0-9]+\]"
    'sets case insensitivity
    objRegExpVisibleCitationItems.ignoreCase = False
    'sets global applicability
    objRegExpVisibleCitationItems.GlobalSearch = True

    strOrphanCitationItems = ""


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

                'checks that the string can be compared
                If objRegExpVisibleCitationItems.Test(Selection) Then
                    'gets the matches (all entries in the citation according to the regular expression)
                    Set colMatchesVisibleCitationItems = objRegExpVisibleCitationItems.Execute(Selection)

                    'treats all matches (all entries in citation) to generate hyperlinks
                    For Each objMatchVisibleCitationItem In colMatchesVisibleCitationItems
                        'weird things happen when there is only one citation item
                        '(the hyperlink is created in the bibliography instead of the citation)
                        'this happens when the full text of the Selection matches the regular expression to find the number of the reference
                        'needs to make sure this does not happen (it only happens with Mendeley Reference Manager 2.x)
                        'somehow the Selection moves form the citation content control to the bibliography content control
                        If colMatchesVisibleCitationItems.Count > 1 Then
                            blnIsSingleCitation = False
                        Else
                            blnIsSingleCitation = True
                        End If

                        'gets the citation number as integer
                        'this will also eliminate leading zeros in numbers (in case of manual modifications)
                        intCitationNumber = CInt(Mid(objMatchVisibleCitationItem.value, 2, Len(objMatchVisibleCitationItem.value) - 2))

                        'to make sure the citation number as text is the same as numeric
                        'and that the citation number is in the bibliography
                        If (("[" & CStr(intCitationNumber) & "]") = objMatchVisibleCitationItem.value) And (intCitationNumber > 0 And intCitationNumber < intReferenceNumber) Then
                            blnCitationNumberFound = True
                        Else
                            blnCitationNumberFound = False
                        End If

                        'if a number of a citation was found (shall always find it), inserts the hyperlink
                        If blnCitationNumberFound Then
                            'according to the version of Mendeley
                            Select Case intAvailableMendeleyVersion
                                'Mendeley Desktop 1.x is installed
                                Case 1
                                    'selects the current field (Mendeley's citation field)
                                    objCurrentFieldOrContentControl.Select
                                    'finds and selects the text of the number of the reference, including square brackets to make sure it is the correct one
                                    With Selection.Find
                                        .Forward = True
                                        .Wrap = wdFindStop
                                        .Text = "[" & CStr(intCitationNumber) & "]"
                                        .Execute
                                        blnReferenceNumberFound = .found
                                    End With
                                'Mendeley Reference Manager 2.x is installed
                                Case 2
                                    'selects the current content control (Mendeley's citation content control)
                                    objCurrentFieldOrContentControl.range.Select
                                    'if only one citation item
                                    If blnIsSingleCitation Then
                                        'finds and selects the text of the number of the reference, without square brackets to avoid the weird behavior
                                        With Selection.Find
                                            .Forward = True
                                            .Wrap = wdFindStop
                                            .Text = CStr(intCitationNumber)
                                            .Execute
                                            blnReferenceNumberFound = .found
                                        End With
                                    'if multiple citation items
                                    Else
                                        'finds and selects the text of the number of the reference, including square brackets to make sure it is the correct one
                                        With Selection.Find
                                            .Forward = True
                                            .Wrap = wdFindStop
                                            .Text = "[" & CStr(intCitationNumber) & "]"
                                            .Execute
                                            blnReferenceNumberFound = .found
                                        End With
                                    End If
                            End Select

                            'if a match was found (it should always find it, but good practice)
                            'selects the correct entry text from the citation field
                            If blnReferenceNumberFound Then
                                'according to the version of Mendeley
                                Select Case intAvailableMendeleyVersion
                                    'Mendeley Desktop 1.x is installed
                                    Case 1
                                        'if the square brackets are not part of the hyperlinks
                                        If Not blnIncludeSquareBracketsInHyperlinks Then
                                            'restricts the selection to only the number
                                            Selection.MoveStart Unit:=wdCharacter, Count:=1
                                            Selection.MoveEnd Unit:=wdCharacter, Count:=-1
                                        End If
                                    'Mendeley Reference Manager 2.x is installed
                                    Case 2
                                        'if it is a single citation item in the citation field or content control
                                        If blnIsSingleCitation Then
                                            'if the square brackets are part of the hyperlinks
                                            'selects the whole citation to include the square brackets
                                            'weird things happen if we extend the selection manually with Selection.Move*
                                            'however, the hyperlink may not be easily removed later
                                            If blnIncludeSquareBracketsInHyperlinks Then
                                                'selects the current content control (Mendeley's citation content control)
                                                objCurrentFieldOrContentControl.range.Select
                                            End If
                                        'if there are multiple citation items in the citation field or content control
                                        Else
                                            'if the square brackets are not part of the hyperlinks
                                            If Not blnIncludeSquareBracketsInHyperlinks Then
                                                'restricts the selection to only the number
                                                Selection.MoveStart Unit:=wdCharacter, Count:=1
                                                Selection.MoveEnd Unit:=wdCharacter, Count:=-1
                                            End If
                                        End If
                                End Select

                                'creates the hyperlink for the current citation entry
                                'a cross-reference is not a good idea, it changes the text in citation (or may delete citation):
                                'better to use normal hyperlink:
                                Selection.Hyperlinks.Add Anchor:=Selection.range, _
                                    Address:="", SubAddress:="GAUG_SignetBibliographie_" & format(CStr(intCitationNumber), "00#"), _
                                    ScreenTip:=""
                            End If
                        Else
                            'if the visible citation item could not be linked to an entry in the bibliography
                            strOrphanCitationItems = strOrphanCitationItems & Trim(objMatchVisibleCitationItem.value) & vbCrLf
                        End If

                        'at this point current citation entry is linked to corresponding reference in bibliography

                    Next 'treats all matches (all entries in citation) to generate hyperlinks

                End If 'checks that the string can be compared

            End If 'if it is a citation
        Next 'checks all fields or content controls
    Next 'documentSection

    'at this point all citations are linked to their corresponding reference in bibliography

    'if orphan citations exist
    If Len(strOrphanCitationItems) > 0 Then
        MsgBox "Orphan citation entries found:" & vbCrLf & vbCrLf & _
            strOrphanCitationItems & vbCrLf & _
            "Remove them from the document or manually create the bookmarks and hyperlinks!", _
            vbExclamation, "GAUG_createHyperlinksForCitationsIEEE()"
    End If

    'if hyperlinks where not generated for the URLs in the bibliography
    If intTotalURLsWithoutHyperlink > 0 Then
        MsgBox "The bibliography contains " & intTotalURLsWithoutHyperlink & " URLs for which" & vbCrLf & _
            "the macros could not create the hyperlinks." & vbCrLf & vbCrLf & _
            "Create the hyperlinks manually by selecting every URL," & vbCrLf & _
            "then go to the menu Insert and click on Link.", _
            vbExclamation, "GAUG_createHyperlinksForCitationsIEEE()"
        End
    End If

    'returns to original position in the document
    currentPosition.Select

    'reenables the screen updating
    Application.ScreenUpdating = True

End Sub



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2024-10-09                                                                                                                **
'**                                                                                                                                           **
'**  Sub GAUG_removeHyperlinksForCitations(Optional ByVal strTypeOfExecution As String = "RemoveHyperlinks")                                  **
'**                                                                                                                                           **
'**  This is an improved version which runs much faster, but still considered as experimental.                                                **
'**  Make sure you have a backup of your document before you execute it!                                                                      **
'**                                                                                                                                           **
'**  Removes the bookmarks generated by GAUG_createHyperlinksForCitations* in the bibliography inserted by Mendeley's plugin.                 **
'**  Removes the hyperlinks generated by GAUG_createHyperlinksForCitations* of the citations inserted by Mendeley's plugin.                   **
'**  Removes the hyperlinks generated by GAUG_createHyperlinksForCitations* in the bibliography inserted by Mendeley's plugin.                **
'**  Removes all manual modifications to the citations and bibliography if specified                                                          **
'**                                                                                                                                           **
'**  Parameter strTypeOfExecution can have three different values:                                                                            **
'**  "RemoveHyperlinks":                                                                                                                      **
'**     UNEXPECTED RESULTS IF MANUAL MODIFICATIONS EXIST, BUT THE FASTEST                                                                     **
'**        Removes the bookmarks and hyperlinks                                                                                               **
'**        Manual modifications to citations and bibliography will remain intact                                                              **
'**  "CleanEnvironment": (only available when Mendeley Desktop 1.x is used)                                                                   **
'**     EXPERIMENTAL, BUT VERY FAST                                                                                                           **
'**        Removes the bookmarks and hyperlinks                                                                                               **
'**        Manual modifications to citations and bibliography will also be removed to have a clean environment                                **
'**  "CleanFullEnvironment": (only available when Mendeley Desktop 1.x is used)                                                               **
'**     SAFE, BUT VERY SLOW IN LONG DOCUMENTS                                                                                                 **
'**        Removes the bookmarks and hyperlinks                                                                                               **
'**        Manual modifications to citations and bibliography will also be removed to have a clean environment                                **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Sub GAUG_removeHyperlinksForCitations(Optional ByVal strTypeOfExecution As String = "RemoveHyperlinks")

    Dim intAvailableMendeleyVersion, intUseMendeleyVersion As Integer
    Dim documentSection As Section
    Dim objCurrentFieldOrContentControl As Object
    Dim fieldBookmark As Bookmark
    Dim selectionHyperlinks As Hyperlinks
    Dim i As Integer
    Dim blnFound, blnCitationFound, blnBibliographyFound As Boolean
    Dim sectionFieldName, sectionFieldNewName As String
    Dim objMendeleyApiClient As Object
    Dim cbbUndoEditButton As CommandBarButton
    Dim blnMabEntwickeltSich As Boolean
    Dim stlStyleInDocument As Word.Style
    Dim strStyleForTitleOfBibliography As String
    Dim blnStyleForTitleOfBibliographyFound As Boolean
    Dim varFieldsOrContentControls As Variant
    Dim currentPosition As range
    'Const INSERT_CITATION_TEXT = "{Formatting Citation}" 'copied from Mendeley Desktop 1.x (from the MS Word plugin Mendeley Cite-O-Matic)
    Const INSERT_CITATION_TEXT = vbNullString


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
    strStyleForTitleOfBibliography = "Titre de dernière section"

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
                    "Cannot continue removing hyperlinks.", _
                    vbCritical, "GAUG_removeHyperlinksForCitations(strTypeOfExecution)"
                'the execution option is not correct
                End
            'if Mendeley Desktop 1.x is used
            Else
                'makes sure that the plugin is installed in Microsoft Word
                If Not GAUG_isMendeleyCiteOMaticPluginInstalled() Then
                    MsgBox "The MS Word plugin Mendeley Cite-O-Matic was not found." & vbCrLf & vbCrLf & _
                        "Install the plugin via Mendeley Desktop 1.x and try again." & vbCrLf & vbCrLf & _
                        "Cannot continue removing hyperlinks.", _
                        vbCritical, "GAUG_removeHyperlinksForCitations(strTypeOfExecution)"
                    'Mendeley's plugin is not installed, cannot call it to remove hyperlinks
                    End
                End If
            End If
            'get the API Client from Mendeley
            Set objMendeleyApiClient = Application.Run("Mendeley.mendeleyApiClient") 'MabEntwickeltSich: This is the way to call the macro directly from Mendeley
        Case "CleanFullEnvironment"
            'only available when Mendeley Desktop 1.x is used
            If Not intAvailableMendeleyVersion = 1 Then
                MsgBox "Incompatible execution type " & Chr(34) & strTypeOfExecution & Chr(34) & " for GAUG_removeHyperlinksForCitations(strTypeOfExecution)." & vbCrLf & vbCrLf & _
                    "Only " & Chr(34) & "RemoveHyperlinks" & Chr(34) & " can be used with Mendeley Reference Manager 2.x (with the App Mendeley Cite)." & vbCrLf & vbCrLf & _
                    "Cannot continue removing hyperlinks.", _
                    vbCritical, "GAUG_removeHyperlinksForCitations(strTypeOfExecution)"
                'the execution option is not correct
                End
            'if Mendeley Desktop 1.x is used
            Else
                'makes sure that the plugin is installed in Microsoft Word
                If Not GAUG_isMendeleyCiteOMaticPluginInstalled() Then
                    MsgBox "The MS Word plugin Mendeley Cite-O-Matic was not found." & vbCrLf & vbCrLf & _
                        "Install the plugin via Mendeley Desktop 1.x and try again." & vbCrLf & vbCrLf & _
                        "Cannot continue removing hyperlinks.", _
                        vbCritical, "GAUG_removeHyperlinksForCitations(strTypeOfExecution)"
                    'Mendeley's plugin is not installed, cannot call it to remove hyperlinks
                    End
                End If
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
                        'hyperlinks are also fields, we can delete them via the more reliable object Field
                        'iterates over all fields of the selection (which include the hyperlinks)
                        For i = Selection.Fields.Count To 1 Step -1
                            'if the field is an hyperlink
                            If Selection.Fields(i).Type = wdFieldHyperlink Then
                                'this is the way to remove hyperlinks without the errors produced by Hyperlink.Delete
                                Selection.Fields(i).Unlink
                            End If
                        Next i
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
                        'checks if it is the bibliography
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

                    'hyperlink are also fields, we can delete them via the more reliable object Field
                    'iterates over all fields of the selection (which include all URL hyperlinks)
                    For i = Selection.Fields.Count To 1 Step -1
                        'if the field is an hyperlink
                        If Selection.Fields(i).Type = wdFieldHyperlink Then
                            'this is the way to remove the URL hyperlinks without the errors produced by Hyperlink.Delete
                            Selection.Fields(i).Unlink
                        End If
                    Next i

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



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2017-01-11                                                                                                                **
'**                                                                                                                                           **
'**  Sub GAUG_removeHyperlinks()                                                                                                              **
'**                                                                                                                                           **
'**  Calls Sub GAUG_removeHyperlinksForCitations(strTypeOfExecution As String) with parameter strTypeOfExecution = "RemoveHyperlinks"         **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Sub GAUG_removeHyperlinks()
    'removes all bookmarks and hyperlinks from the citations and bibliography
    Call GAUG_removeHyperlinksForCitations("RemoveHyperlinks")
End Sub



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2017-01-11                                                                                                                **
'**                                                                                                                                           **
'**  Sub GAUG_cleanEnvironment()                                                                                                              **
'**                                                                                                                                           **
'**  Calls Sub GAUG_removeHyperlinksForCitations(strTypeOfExecution As String) with parameter strTypeOfExecution = "CleanEnvironment"         **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Sub GAUG_cleanEnvironment()
    'removes all bookmarks, hyperlinks and manual modifications to the citations and bibliography
    Call GAUG_removeHyperlinksForCitations("CleanEnvironment")
End Sub



'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
'**  Author: José Luis González García                                                                                                        **
'**  Last modified: 2017-01-11                                                                                                                **
'**                                                                                                                                           **
'**  Sub GAUG_cleanFullEnvironment()                                                                                                          **
'**                                                                                                                                           **
'**  Calls Sub GAUG_removeHyperlinksForCitations(strTypeOfExecution As String) with parameter strTypeOfExecution = "CleanFullEnvironment"     **
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Sub GAUG_cleanFullEnvironment()
    'removes all bookmarks, hyperlinks and manual modifications to the citations and bibliography
    Call GAUG_removeHyperlinksForCitations("CleanFullEnvironment")
End Sub



