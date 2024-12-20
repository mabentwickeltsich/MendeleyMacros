# Mendeley Macros 2.0
***Now available for Windows and macOS!***

---

Macros used with Microsoft Word to add new functionalities to Mendeley's plugin. They support Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic) and Mendeley Reference Manager 2.x (with the App Mendeley Cite).

**Functionalities:**
* The macros `GAUG_*` generate hyperlinks for citations pointing to the corresponding entry in the bibliography, as well as hyperlinks for the URLs in the bibliography generated by Mendeley's plugin. They support IEEE and APA CSL citation styles.
* The modified macro `refreshDocument` maintains Microsoft Word style of the bibliography generated by Mendeley's plugin when refreshing it. It is only available for Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic).

*If you find the macros useful and would like to show your gratitude, please consider making a donation to help keep the project alive: [PayPal.Me](https://paypal.me/MabEntwickeltSich "Donate to MabEntwickeltSich")*

## Author
José Luis González García

## Overview
You need all the files for your platform to create or remove the hyperlinks. Just import them to Microsoft Word and execute any of the macros `GAUG_*`. macOS requires a regular expression engine which can be downloaded [here](https://github.com/mabentwickeltsich/vba-regex/blob/65f697a3c499a4cd2d24e288dc60a9070f0e2517/aio/build/StaticRegexSingle.bas). The default custom configuration allows you to use the macros without any further modification. See the new detailed instructions bellow on how to install them.
Alternatively, you can follow the *Quick start guides* for Windows and macOS available on [my YouTube channel](https://www.youtube.com/@MabEntwickeltSich "MabEntwickeltSich on YouTube").

For those who still use Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic). If you are annoyed that Mendeley's plugin changes the font style of the bibliography every time you refresh it, then you need to modify the macro `refreshDocument` installed by Mendeley's plugin to correct the problem. If you are happy with your bibliography, just forget about `refreshDocument`. See the detailed instructions bellow on how to modify the original macro.

## Macros

#### GAUG_removeHyperlinksForCitations(strTypeOfExecution)
Removes from the document all hyperlinks generated for the citations and leaves all fields as originally inserted by Mendeley's plugin. It also removes all bookmarks and hyperlinks generated in the bibliography.

The parameter `strTypeOfExecution` can have one of three different values:
* `"RemoveHyperlinks"`: Removes the bookmarks and hyperlinks from the document. Manual modifications to citations and bibliography will remain intact. WARNING: UNEXPECTED RESULTS IF MANUAL MODIFICATIONS EXIST, BUT VERY FAST.
* `"CleanEnvironment"`: Removes the bookmarks and hyperlinks from the document. Manual modifications to citations and bibliography will also be removed to have a clean environment. This value CAN ONLY be used with Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic). WARNING: STILL EXPERIMENTAL, BUT FAST.
* `"CleanFullEnvironment"`: Removes the bookmarks and hyperlinks from the document. Manual modifications to citations and bibliography will also be removed to have a clean environment. This value CAN ONLY be used with Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic). WARNING: SAFE, BUT VERY SLOW IN LONG DOCUMENTS.

Custom configuration:
* `intUseMendeleyVersion`: The default value `0` indicates that the macro should automatically detect the version of Mendeley; IT IS THE WAY TO GO. If, for some reason, the macro could not detect the version of Mendeley, it can be set manually. The value `1` forces the macro to use Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic) and the value `2` forces the macro to use Mendeley Reference Manager 2.x (with the App Mendeley Cite).
* `strStyleForTitleOfBibliography`: Specifies the Microsoft Word style used for the title of the bibliography. It helps to improve speed when locating the bibliography by narrowing the search to only the sections of the document with this style. It is only used when the flag `blnMabEntwickeltSich` is set to `True`. The document MUST follow a particular structure; see the description of the flag `blnMabEntwickeltSich`. The default value is a custom style called “Titre de dernière section” that you may need to create or change.
* `blnMabEntwickeltSich`: The default value `False` will force the macro to check every section of the document for the bibliography; IT IS THE WAY TO GO. If the flag is set to `True`, the speed may improve in long documents, but you need a particular structure in your document. First, the style specified by `strStyleForTitleOfBibliography` MUST exist otherwise the macro will throw an error. In addition to the style, the bibliography MUST also be placed in a section with a title in that style or the macro will not find it.

**Requires:**
* `GAUG_isMendeleyCiteOMaticPluginInstalled()`
* `GAUG_getAvailableMendeleyVersion(intUseMendeleyVersion)`

**Important:** It works ONLY with the IEEE and APA CSL citation styles installed with Mendeley. It MAY not work if you make manual modifications to the citation fields, bibliography or both IEEE and APA CSL citation styles.

**NOTE:** `strTypeOfExecution="RemoveHyperlinks"` is the way to go; it is reliable and fast. In case of manual modifications to the citations or bibliography, the hyperlinks may not be removed or unlinked correctly. If previous approach did not work, use `strTypeOfExecution="CleanEnvironment"`; also fast but still considered as EXPERIMENTAL. It executes a very light weight version of Mendeley's function `Mendeley.undoEdit` to remove the hyperlinks and manual modifications from the citations. If you want to be in the safe side, use `strTypeOfExecution="CleanFullEnvironment"` which calls Mendeley's functions to do the job. The execution with the last option may take several minutes; use it with caution. Mendeley is slow when undoing changes to citation fields. The last two options are only available with Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic). This macro is also called when creating hyperlinks for IEEE and APA citation styles.

#### GAUG_removeHyperlinks()
Wrapper for `GAUG_removeHyperlinksForCitations("RemoveHyperlinks")`.

#### GAUG_cleanEnvironment()
Wrapper for `GAUG_removeHyperlinksForCitations("CleanEnvironment")`.

#### GAUG_cleanFullEnvironment()
Wrapper for `GAUG_removeHyperlinksForCitations("CleanFullEnvironment")`.

#### GAUG_createHyperlinksForCitationsAPA()
Creates the bookmarks in the bibliography and the hyperlinks for the citations in the document. It also creates the hyperlinks for the URLs in the bibliography. The citations must follow the APA CSL citation style.

Custom configuration:
* `intUseMendeleyVersion`: The default value `0` indicates that the macro should automatically detect the version of Mendeley; IT IS THE WAY TO GO. If, for some reason, the macro could not detect the version of Mendeley, it can be set manually. The value `1` forces the macro to use Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic) and the value `2` forces the macro to use Mendeley Reference Manager 2.x (with the App Mendeley Cite).
* `strStyleForTitleOfBibliography`: Specifies the Microsoft Word style used for the title of the bibliography. It helps to improve speed when locating the bibliography by narrowing the search to only the sections of the document with this style. It is only used when the flag `blnMabEntwickeltSich` is set to `True`. The document MUST follow a particular structure; see the description of the flag `blnMabEntwickeltSich`. The default value is a custom style called “Titre de dernière section” that you may need to create or change.
* `blnMabEntwickeltSich`: The default value `False` will force the macro to check every section of the document for the bibliography; IT IS THE WAY TO GO. If the flag is set to `True`, the speed may improve in long documents, but you need a particular structure in your document. First, the style specified by `strStyleForTitleOfBibliography` MUST exist otherwise the macro will throw an error. In addition to the style, the bibliography MUST also be placed in a section with a title in that style or the macro will not find it.
* `blnGenerateHyperlinksForURLs`: The default value `True` will generate the hyperlinks for the URLs in the bibliography. The regular expression used to detect the URLs may not find them all; in such case you need to manually specify those that were not detected or fully detected. If the flag is set to `False`, the macro will not generate the hyperlinks for the URLs in the bibliography.
* `arrNonDetectedURLs`: Specifies the URLs that were not detected, or fully detected, in the bibliography by the regular expression. The macro will use these URLs to generate the corresponding hyperlinks. The list is only used when the flag `blnGenerateHyperlinksForURLs` is set to `True`.
* `strTypeOfExecution`: The default value `"RemoveHyperlinks"` is the fastest and THE WAY TO GO. See more details in the description of the macro `GAUG_removeHyperlinksForCitations(strTypeOfExecution)` about the type of execution.

**Requires:**
* `GAUG_getAvailableMendeleyVersion(intUseMendeleyVersion)`
* `GAUG_getAllCitationsFullInformation(intMendeleyVersion)`
* `GAUG_removeHyperlinksForCitations(strTypeOfExecution)`
* `GAUG_getCitationFullInfo(intMendeleyVersion, strAllCitationsFullInformation, fldCitation, ccCitation)`
* `GAUG_getCitationItemsFromCitationFullInfo(intMendeleyVersion, strCitationFullInfo)`
* `GAUG_getPartsFromVisibleCitationItem(strVisibleCitationItem)`
* `GAUG_getAuthorsEditorsFromCitationItem(intMendeleyVersion, strAuthorEditor, strCitationItem)`
* `GAUG_getYearFromCitationItem(intMendeleyVersion, strCitationItem)`
* `GAUG_getSafeStringForRegularExpressions(strOriginalString)`
* `GAUG_createHyperlinksForURLsInBibliography(intMendeleyVersion, fldBibliography, ccBibliography, arrNonDetectedURLs())`

**Important:** It works ONLY with the APA CSL citation style installed with Mendeley. It MAY not work if you make manual modifications to the citation fields, bibliography or APA CSL citation style.

#### GAUG_createHyperlinksForCitationsIEEE()
Creates the bookmarks in the bibliography and the hyperlinks for the citations in the document. It also creates the hyperlinks for the URLs in the bibliography. The citations must follow the IEEE CSL citation style.

Custom configuration:
* `intUseMendeleyVersion`: The default value `0` indicates that the macro should automatically detect the version of Mendeley; IT IS THE WAY TO GO. If, for some reason, the macro could not detect the version of Mendeley, it can be set manually. The value `1` forces the macro to use Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic) and the value `2` forces the macro to use Mendeley Reference Manager 2.x (with the App Mendeley Cite).
* `strStyleForTitleOfBibliography`: Specifies the Microsoft Word style used for the title of the bibliography. It helps to improve speed when locating the bibliography by narrowing the search to only the sections of the document with this style. It is only used when the flag `blnMabEntwickeltSich` is set to `True`. The document MUST follow a particular structure; see the description of the flag `blnMabEntwickeltSich`. The default value is a custom style called “Titre de dernière section” that you may need to create or change.
* `blnMabEntwickeltSich`: The default value `False` will force the macro to check every section of the document for the bibliography; IT IS THE WAY TO GO. If the flag is set to `True`, the speed may improve in long documents, but you need a particular structure in your document. First, the style specified by `strStyleForTitleOfBibliography` MUST exist otherwise the macro will throw an error. In addition to the style, the bibliography MUST also be placed in a section with a title in that style or the macro will not find it.
* `blnGenerateHyperlinksForURLs`: The default value `True` will generate the hyperlinks for the URLs in the bibliography. The regular expression used to detect the URLs may not find them all; in such case you need to manually specify those that were not detected or fully detected. If the flag is set to `False`, the macro will not generate the hyperlinks for the URLs in the bibliography.
* `arrNonDetectedURLs`: Specifies the URLs that were not detected, or fully detected, in the bibliography by the regular expression. The macro will use these URLs to generate the corresponding hyperlinks. The list is only used when the flag `blnGenerateHyperlinksForURLs` is set to `True`.
* `blnIncludeSquareBracketsInHyperlinks`: The default value `False` will exclude the square brackets from the hyperlinks. If the flag is set to `True`, the macro will include the square brackets as part of the hyperlinks.
* `strTypeOfExecution`: The default value `"RemoveHyperlinks"` is the fastest and THE WAY TO GO. See more details in the description of the macro `GAUG_removeHyperlinksForCitations(strTypeOfExecution)` about the type of execution.

**Requires:**
* `GAUG_getAvailableMendeleyVersion(intUseMendeleyVersion)`
* `GAUG_removeHyperlinksForCitations(strTypeOfExecution)`
* `GAUG_createHyperlinksForURLsInBibliography(intMendeleyVersion, fldBibliography, ccBibliography, arrNonDetectedURLs())`

**Important:** It works ONLY with the IEEE CSL citation style installed with Mendeley. It MAY not work if you make manual modifications to the citation fields, bibliography or IEEE CSL citation style.

#### The rest of GAUG_*()
Helper functions and classes for `GAUG_removeHyperlinksForCitations(strTypeOfExecution)`, `GAUG_createHyperlinksForCitationsAPA()` and `GAUG_createHyperlinksForCitationsIEEE()`

#### refreshDocument
Only for those who still use Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic). This is a copy of the original macro installed by Mendeley and located in *Mendeley-1.16.1.dotm* (file name depends on the version of Mendeley). It has been modified in order to keep the Microsoft Word style of the bibliography generated by Mendeley's plugin when refreshing it. Have a look at the three lines with the comment `'MabEntwickeltSich`, you need to add them to the macro of your installation of Mendeley. Follow the installation instructions below.

Keep in mind that this will only work if you used a Microsoft Word style for the bibliography. You can use a build-in style or create your own. If you made manual modifications to the format (font and paragraph) directly to the bibliography and then refresh it, it will go back to the original Microsoft Word style settings.

**Important:** This macro will NOT work out if its context. DO NOT COPY THE MACRO TO MICROSOFT WORD. This is just to illustrate what modifications have to be done to the original macro installed by Mendeley's plugin.

## Installation

Remove any previous installation of Mendeley Macros before installing the most recent release; it prevents conflicts and makes your life easier.

#### GAUG_*
1. Open Microsoft Visual Basic for Applications from the “Developer” tab in Microsoft Word. You may have to enable the “Developer” tab if you do not see it.
For more information on how to show the “Developer” tab, check Microsoft Office support:
[https://support.office.com/en-us/article/Show-the-Developer-tab-E1192344-5E56-4D45-931B-E5FD9BEA2D45](https://support.office.com/en-us/article/Show-the-Developer-tab-E1192344-5E56-4D45-931B-E5FD9BEA2D45 "Show the Developer tab")

2. Import all the files, that correspond to your platform, to your Microsoft Word to install the macros `GAUG_*`.
For more information about macros, check Microsoft Office support:
[https://support.office.com/en-us/article/Create-or-run-a-macro-C6B99036-905C-49A6-818A-DFB98B7C3C9C](https://support.office.com/en-us/article/Create-or-run-a-macro-C6B99036-905C-49A6-818A-DFB98B7C3C9C "Create or run a macro")

3. The macros GAUG_* make use of regular expressions; hence you need to enable/install the regular expression engine.

    a. On Windows: This step is optional; only required if the macros ask for it. Enable the RegExp object in Microsoft Visual Basic for Applications. Open the menu “Tools” | “References” and check the box next to “Microsoft VBScript Regular Expressions 5.5”.
    For more information on how to activate the RegExp object, check Microsoft Office VBA Reference:
    [https://learn.microsoft.com/en-us/office/vba/Language/How-to/check-or-add-an-object-library-reference](https://learn.microsoft.com/en-us/office/vba/Language/How-to/check-or-add-an-object-library-reference "Check or add an object library reference").

    b. On macOS: You NEED the regular expression engine [StaticRegexSingle](https://github.com/mabentwickeltsich/vba-regex/blob/65f697a3c499a4cd2d24e288dc60a9070f0e2517/aio/build/StaticRegexSingle.bas); download it and import it to your Microsoft Word. It is a fork of another project that has been adapted for Mendeley Macros.

4. The macros `GAUG_*` make use of Streams on Windows and AppleScripts on macOS; hence you need to enable/install them. The Streams and AppleScripts are ONLY used with large documents; however, they need to be enabled/installed to prevent errors.

    a. On Windows: This step is optional; only required if the macros ask for it. Enable the ADODB object in Microsoft Visual Basic for Applications. Open the menu “Tools” | “References” and check the box next to “Microsoft ActiveX Data Objects 6.1 Library”.
    For more information on how to activate the ADODB object, check Microsoft Office VBA Reference:
    [https://learn.microsoft.com/en-us/office/vba/Language/How-to/check-or-add-an-object-library-reference](https://learn.microsoft.com/en-us/office/vba/Language/How-to/check-or-add-an-object-library-reference "Check or add an object library reference")

    b. On macOS: Download the file `MendeleyMacrosHelper_Mac.scpt` and copy it to *~/Library/Application Scripts/com.microsoft.Word/*. You can access the Library by typing `open ~/Library` on the terminal.

#### refreshDocument
Only for those who still use Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic). To “install” `refreshDocument` you need to modify the original macro by inserting the three extra lines with the comment `'MabEntwickeltSich`:

1. Open the file *Mendeley-1.16.1.dotm* (or *Mendeley-word201x-1.16.1.dot* on macOS) located in *C:\Program Files\Mendeley Desktop\wordPlugin* (or */Applications/Mendeley Desktop.app/Contents/Resources/macWordPlugin/word201x* on macOS). Check your own installation, your version of Mendeley may be different. You may need to execute Microsoft Word as administrator to be able to save the changes.
You may have to enable macros for the document when you open it.

2. Open Microsoft Visual Basic for Applications from the “Developer” tab in Microsoft Word. You may have to enable the “Developer” tab if you do not see it.
For more information on how to show the “Developer” tab, check Microsoft Office support:
[https://support.office.com/en-us/article/Show-the-Developer-tab-E1192344-5E56-4D45-931B-E5FD9BEA2D45](https://support.office.com/en-us/article/Show-the-Developer-tab-E1192344-5E56-4D45-931B-E5FD9BEA2D45 "Show the Developer tab")

3. Modify the macro:
`Function refreshDocument(Optional openingDocument As Boolean = False) As Boolean`
which is located in the module *MendeleyLib*.

4. Once you finish the modifications save the changes, close Microsoft Visual Basic for Applications and Microsoft Word.

5. Open Mendeley Desktop 1.x, uninstall the Microsoft Word plugin and reinstall it again for the changes to take effect.

## Usage
ALWAYS have a BACKUP COPY of your document BEFORE using these macros.

Execute the desired macro. See also **Extending/modifying the code**.

When your document is large enough (around 180 or more citations), Mendeley Macros will extract the contents of the document into a temporary folder in order to read the full information of all the citations. This is the only way due to an error within Microsoft Word which prevents Mendeley Macros to directly gather the information. On Windows, this process is transparent, you will not notice it. On macOS, due to the sandbox restrictions, Mendeley Macros make use of the AppleScripts described on the installation. You MUST grant permissions to be able to access the folder where your document is located, the temporary folder (that Mendeley Macros create) and the file webextension1.xml (extracted from your document into the temporary folder).

My recommendation for a typical usage is to keep your document free of any manual modification to the bibliography or to the citations inserted by Mendeley, but you can merge the citations with the standard way provided by Mendeley: [1][2][3][4] becomes [1]-[4]. This also applies for the APA CSL citation style.

Also, keep your document without the bookmarks and hyperlinks generated by these macros while editing. Generate them ONLY when you want to create the PDF file or when you are done with the edition. If you want to further edit the document, remove all bookmarks and hyperlinks generated by these macros to prevent any conflict with Mendeley's plugin; the citation numbers in IEEE or text in APA may change.

Only for those who still use Mendeley Desktop 1.x (with the MS Word plugin Mendeley Cite-O-Matic). It is important to note that `GAUG_removeHyperlinksForCitations(strTypeOfExecution)` is very slow when `strTypeOfExecution="CleanFullEnvironment"`. It uses Mendeley's code to restore the original citation fields and bibliography. It is also called from `GAUG_createHyperlinksForCitationsAPA()` and `GAUG_createHyperlinksForCitationsIEEE()` to have a clean environment before creating the bookmarks and hyperlinks.

## Extending/modifying the code
There is no need to do any changes to the code to start using the macros. The default custom configuration, at the beginning of every macro, fits most scenarios. The flag `blnMabEntwickeltSich` activates part of the code to improve speed in long documents, but a particular document structure is required or errors will appear during the execution.

In this moment the code is adapted to my own needs and to the structure of my document. Nevertheless, changing the code to fit other requirements is straight forward when you stick to the IEEE or APA CSL citation styles. Much more effort may be required to support a different CLS citation style; and I can confirm that great effort was done to include APA (even more to support the new version of Mendeley!).

My document is divided in sections (Microsoft Word sections) for the chapters and other parts that are included in the thesis. The bibliography is located in a section with the title “Bibliography” using the style “Titre de dernière section” (custom Microsoft Word style for the title). The macros to remove or create the hyperlinks will try to find the bibliography in a section with this description when the flag `blnMabEntwickeltSich` is set to `True`. I did it this way to increase speed in long documents. If you want to remove this restriction and locate the bibliography in any section, simply set the flag to `False` (the default value) to force the macros to check every section for the bibliography.

Comments and suggestions are welcome.
