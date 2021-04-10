Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports System.IO
Imports System.Web
Imports System.Data
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Module Main
    Const ArrayDim = 18                             ' Number of Columns/Dimentions in the scraped mail metada array
    Const FileLocation As String = "Documents"      ' [UNUSED | PLACE HOLDER] to be used in lue of a flolder picker
    Public Selected_mail_items(,) As String          ' An object that sores all the metada of selected mail items
    Private ext As String                           ' Used to store the file extention of an exported json/xml file
    Private exportString As String                  ' This is the final string of text to be written to a file
    Private filePathPicked As String                ' This is the selected folder location stored as a string/text

    Public Sub save_EmailsWithAttments()
        Dim i As Integer
        Dim myOlApp As Outlook.Application = New Outlook.Application
        Dim objView As Outlook.Explorer = myOlApp.ActiveExplorer
        Dim oMail As Outlook.MailItem
        Dim MailMetadata As Array
        Dim olAttachment As Outlook.Attachment
        Dim TextFile As Integer
        Dim FilePath As String
        Dim FileName As String
        Dim FilePathConverter As String
        Dim file As System.IO.StreamWriter

        Const olMsg As Long = 0
        Dim path As String = FolderPicker()

        For Each olMail In objView.Selection
            FileName = olMail.Subject
            FileName = Replace(FileName, "\", " ")
            FileName = Replace(FileName, "/", " ")
            FileName = Replace(FileName, ".", " ")
            FileName = Replace(FileName, "|", " ")
            FileName = Replace(FileName, "*", " ")
            FileName = Replace(FileName, "*", " ")
            FileName = Replace(FileName, "?", " ")
            FileName = Replace(FileName, ":", " ")
            FileName = Replace(FileName, "<", " ")
            FileName = Replace(FileName, ">", " ")
            Dim savepath As String = path & "\" & FileName & ".txt"
            olMail.saveas(savepath, olMsg)

        Next


        On Error Resume Next
        'filePathPicked = FolderPicker()
        'MkDir "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\"
        exportString = ""
        For Each olMail In objView.Selection


            'Extract all Attachments and place into their own folder
            'the folder name matched the wmail item json name
            On Error Resume Next
            For Each olAttachment In olMail.Attachments
                FileName = olMail.Subject
                FileName = Replace(FileName, "\", " ")
                FileName = Replace(FileName, "/", " ")
                FileName = Replace(FileName, ".", " ")
                FileName = Replace(FileName, "|", " ")
                FileName = Replace(FileName, "*", " ")
                FileName = Replace(FileName, "*", " ")
                FileName = Replace(FileName, "?", " ")
                FileName = Replace(FileName, ":", " ")
                FileName = Replace(FileName, "<", " ")
                FileName = Replace(FileName, ">", " ")
                If olAttachment.FileName <> "" Then
                    'MkDir "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\" & FileName & "\"
                    MkDir(path & "\" & FileName & "\")
                    'FilePathConverter = File_Exists("C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\" & FileName & "\" & olAttachment.FileName)
                    FilePathConverter = File_Exists(path & "\" & FileName & "\" & olAttachment.FileName)
                    olAttachment.SaveAsFile(FilePathConverter)
                End If
            Next olAttachment
        Next olMail

    End Sub

    Public Sub save_Emails()
        Const olMsg As Long = 0
        Dim myOlApp As Outlook.Application = New Outlook.Application
        Dim objView As Outlook.Explorer = myOlApp.ActiveExplorer
        Dim oMail As Outlook.MailItem
        Dim path As String = FolderPicker()

        For Each olMail In objView.Selection
            Dim FileName As String = olMail.Subject
            FileName = Replace(FileName, "\", " ")
            FileName = Replace(FileName, "/", " ")
            FileName = Replace(FileName, ".", " ")
            FileName = Replace(FileName, "|", " ")
            FileName = Replace(FileName, "*", " ")
            FileName = Replace(FileName, "*", " ")
            FileName = Replace(FileName, "?", " ")
            FileName = Replace(FileName, ":", " ")
            FileName = Replace(FileName, "<", " ")
            FileName = Replace(FileName, ">", " ")
            Dim savepath As String = path & "\" & FileName & ".txt"
            olMail.saveas(savepath, olMsg)

        Next

    End Sub

    Public Sub JSON()
        Dim i As Integer
        Dim myOlApp As Outlook.Application = New Outlook.Application
        Dim objView As Outlook.Explorer = myOlApp.ActiveExplorer
        Dim oMail As Outlook.MailItem
        Dim MailMetadata As Array
        Dim olAttachment As Outlook.Attachment
        Dim TextFile As Integer
        Dim FilePath As String
        Dim FileName As String
        Dim FilePathConverter As String
        Dim file As System.IO.StreamWriter

        On Error Resume Next
        filePathPicked = FolderPicker()
        'MkDir "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\"
        exportString = ""
        For Each olMail In objView.Selection
            Dim jsonArrays As Collection
            jsonArrays = New Collection
            'Creating json sub Arrays
            jsonArrays.Add(Item:=jsonArray(olMail.To, ";"))
            jsonArrays.Add(Item:=jsonArray(olMail.CC, ";"))
            'Creating the main json array
            exportString = "{" & vbNewLine & vbTab &
                            """people"" : {" & vbNewLine & vbTab & vbTab &
                                """to"" : " & jsonArrays(1) & "," & vbNewLine & vbTab & vbTab &
                                """cc"" : " & jsonArrays(2) & vbNewLine & vbTab &
                            "}," & vbNewLine & vbTab
            exportString = exportString &
                            """names"" : {" & vbNewLine & vbTab & vbTab &
                                """ReplyRecipientNames"" : """ & olMail.ReplyRecipientNames & """," & vbNewLine & vbTab & vbTab &
                                """SenderName"" : """ & olMail.SenderName & """," & vbNewLine & vbTab & vbTab &
                                """SentOnBehalfOfName"" : """ & olMail.SentOnBehalfOfName & """," & vbNewLine & vbTab & vbTab &
                                """ReceivedOnBehalfOfName"" : """ & olMail.ReceivedOnBehalfOfName & """," & vbNewLine & vbTab & vbTab &
                                """ReceivedByName"" : """ & olMail.ReceivedByName & """" & vbNewLine & vbTab &
                            "}," & vbNewLine & vbTab
            exportString = exportString &
                            """time"" : {" & vbNewLine & vbTab & vbTab &
                                """CreationTime"" : """ & olMail.CreationTime & """," & vbNewLine & vbTab & vbTab &
                                """LastModificationTime"" : """ & olMail.LastModificationTime & """," & vbNewLine & vbTab & vbTab &
                                """SentOn"" : """ & olMail.SentOn & """," & vbNewLine & vbTab & vbTab &
                                """ReceivedTime"" : """ & olMail.ReceivedTime & """" & vbNewLine & vbTab &
                            "}," & vbNewLine & vbTab
            exportString = exportString &
                            """metadata"" : {" & vbNewLine & vbTab & vbTab &
                                """SenderEmailType"" : """ & olMail.SenderEmailType & """," & vbNewLine & vbTab & vbTab &
                                """Size"" : " & olMail.Size & "," & vbNewLine & vbTab & vbTab &
                                """UnRead"" : " & olMail.UnRead & "," & vbNewLine & vbTab & vbTab &
                                """Sent"" : " & olMail.Sent & "," & vbNewLine & vbTab & vbTab &
                                """Importance"" : " & olMail.Importance & vbNewLine & vbTab &
                            "}," & vbNewLine & vbTab
            exportString = exportString &
                            """text"" : {" & vbNewLine & vbTab & vbTab &
                                    """Subject"" : """ & Replace(olMail.Subject, """", "'") & """," & vbNewLine & vbTab & vbTab &
                                    """Body"" : """ & Replace(olMail.Body, """", "'") & """" & vbNewLine & vbTab &
                                "}" & vbNewLine &
                        "}"
            'Create File name
            FileName = Format(olMail.SentOn, "yymmdd") & "-" & Format(olMail.ReceivedTime, "hhmmss") & "-" & olMail.SenderName & "-" & Left(olMail.Subject, 30)
            'Remove reserved characters fron teh file name
            FileName = Replace(FileName, "\", " ")
            FileName = Replace(FileName, "/", " ")
            FileName = Replace(FileName, ".", " ")
            FileName = Replace(FileName, "|", " ")
            FileName = Replace(FileName, "*", " ")
            FileName = Replace(FileName, "*", " ")
            FileName = Replace(FileName, "?", " ")
            FileName = Replace(FileName, ":", " ")
            FileName = Replace(FileName, "<", " ")
            FileName = Replace(FileName, ">", " ")
            'Set the file path
            'FilePath = "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\" & FileName & ".json"
            FilePath = filePathPicked & "\" & FileName & ".json"
            'Insure the file path is unique
            FilePath = File_Exists(FilePath)
            'Write text file (.json)
            TextFile = FreeFile()

            file = My.Computer.FileSystem.OpenTextFileWriter(FilePath, True)
            On Error Resume Next
            file.WriteLine(exportString)
            file.Close()

            'Extract all Attachments and place into their own folder
            'the folder name matched the wmail item json name
            On Error Resume Next
            For Each olAttachment In olMail.Attachments
                If olAttachment.FileName <> "" Then
                    'MkDir "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\" & FileName & "\"
                    MkDir(filePathPicked & "\" & FileName & "\")
                    'FilePathConverter = File_Exists("C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\" & FileName & "\" & olAttachment.FileName)
                    FilePathConverter = File_Exists(filePathPicked & "\" & FileName & "\" & olAttachment.FileName)
                    olAttachment.SaveAsFile(FilePathConverter)
                End If
            Next olAttachment




        Next olMail
    End Sub

    Public Sub JSONPlane()
        Dim i As Integer
        Dim myOlApp As Outlook.Application = New Outlook.Application
        Dim objView As Outlook.Explorer = myOlApp.ActiveExplorer
        Dim oMail As Outlook.MailItem
        Dim MailMetadata As Array
        Dim olAttachment As Outlook.Attachment
        Dim TextFile As Integer
        Dim FilePath As String
        Dim FileName As String
        Dim FilePathConverter As String
        Dim file As System.IO.StreamWriter

        On Error Resume Next
        filePathPicked = FolderPicker()
        'MkDir "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\"
        exportString = ""
        For Each olMail In objView.Selection
            Dim jsonArrays As Collection
            jsonArrays = New Collection
            'Creating json sub Arrays
            jsonArrays.Add(Item:=jsonArray(olMail.To, ";"))
            jsonArrays.Add(Item:=jsonArray(olMail.CC, ";"))
            'Creating the main json array
            exportString = "{" & vbNewLine & vbTab &
                            """people"" : {" & vbNewLine & vbTab & vbTab &
                                """to"" : " & jsonArrays(1) & "," & vbNewLine & vbTab & vbTab &
                                """cc"" : " & jsonArrays(2) & vbNewLine & vbTab &
                            "}," & vbNewLine & vbTab
            exportString = exportString &
                            """names"" : {" & vbNewLine & vbTab & vbTab &
                                """ReplyRecipientNames"" : """ & olMail.ReplyRecipientNames & """," & vbNewLine & vbTab & vbTab &
                                """SenderName"" : """ & olMail.SenderName & """," & vbNewLine & vbTab & vbTab &
                                """SentOnBehalfOfName"" : """ & olMail.SentOnBehalfOfName & """," & vbNewLine & vbTab & vbTab &
                                """ReceivedOnBehalfOfName"" : """ & olMail.ReceivedOnBehalfOfName & """," & vbNewLine & vbTab & vbTab &
                                """ReceivedByName"" : """ & olMail.ReceivedByName & """" & vbNewLine & vbTab &
                            "}," & vbNewLine & vbTab
            exportString = exportString &
                            """time"" : {" & vbNewLine & vbTab & vbTab &
                                """CreationTime"" : """ & olMail.CreationTime & """," & vbNewLine & vbTab & vbTab &
                                """LastModificationTime"" : """ & olMail.LastModificationTime & """," & vbNewLine & vbTab & vbTab &
                                """SentOn"" : """ & olMail.SentOn & """," & vbNewLine & vbTab & vbTab &
                                """ReceivedTime"" : """ & olMail.ReceivedTime & """" & vbNewLine & vbTab &
                            "}," & vbNewLine & vbTab
            exportString = exportString &
                            """metadata"" : {" & vbNewLine & vbTab & vbTab &
                                """SenderEmailType"" : """ & olMail.SenderEmailType & """," & vbNewLine & vbTab & vbTab &
                                """Size"" : " & olMail.Size & "," & vbNewLine & vbTab & vbTab &
                                """UnRead"" : " & olMail.UnRead & "," & vbNewLine & vbTab & vbTab &
                                """Sent"" : " & olMail.Sent & "," & vbNewLine & vbTab & vbTab &
                                """Importance"" : " & olMail.Importance & vbNewLine & vbTab &
                            "}," & vbNewLine & vbTab
            exportString = exportString &
                            """text"" : {" & vbNewLine & vbTab & vbTab &
                                    """Subject"" : """ & Replace(olMail.Subject, """", "'") & """," & vbNewLine & vbTab & vbTab &
                                    """Body"" : """ & Replace(olMail.Body, """", "'") & """" & vbNewLine & vbTab &
                                "}" & vbNewLine &
                        "}"
            'Create File name
            FileName = Format(olMail.SentOn, "yymmdd") & "-" & Format(olMail.ReceivedTime, "hhmmss") & "-" & olMail.SenderName & "-" & Left(olMail.Subject, 30)
            'Remove reserved characters fron teh file name
            FileName = Replace(FileName, "\", " ")
            FileName = Replace(FileName, "/", " ")
            FileName = Replace(FileName, ".", " ")
            FileName = Replace(FileName, "|", " ")
            FileName = Replace(FileName, "*", " ")
            FileName = Replace(FileName, "*", " ")
            FileName = Replace(FileName, "?", " ")
            FileName = Replace(FileName, ":", " ")
            FileName = Replace(FileName, "<", " ")
            FileName = Replace(FileName, ">", " ")
            'Set the file path
            'FilePath = "C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\" & FileName & ".json"
            FilePath = filePathPicked & "\" & FileName & ".json"
            'Insure the file path is unique
            FilePath = File_Exists(FilePath)
            'Write text file (.json)
            TextFile = FreeFile()

            file = My.Computer.FileSystem.OpenTextFileWriter(FilePath, True)
            file.WriteLine(exportString)
            file.Close()

        Next olMail
    End Sub

    Public Sub CSV()
        'https://social.msdn.microsoft.com/Forums/vstudio/en-US/200b3bd7-5328-4218-a0dc-5aaa230908f2/two-dimensional-array-to-datatable?forum=netfxbcl
        Dim table As New DataTable
        table.Columns.Add("To")
        table.Columns.Add("CC")
        table.Columns.Add("Reply_Recipient_Names")
        table.Columns.Add("Sender_Email_Address")
        table.Columns.Add("Sender_Name")
        table.Columns.Add("Sent_On_Behalf_Of_Name")
        table.Columns.Add("Sender_Email_Type")
        table.Columns.Add("Sent")
        table.Columns.Add("Size")
        table.Columns.Add("Unread")
        table.Columns.Add("Creation_Time")
        table.Columns.Add("Last_Modification_Time")
        table.Columns.Add("Sent_On")
        table.Columns.Add("Received_Time")
        table.Columns.Add("Importance")
        table.Columns.Add("Received_By_Name")
        table.Columns.Add("Received_On_Behalf_Of_Name")
        table.Columns.Add("Subject")
        table.Columns.Add("Body")

        Call Mail_Scrape()

        For outerIndex As Integer = 0 To UBound(Selected_mail_items)
            Dim newRow As DataRow = table.NewRow()
            For innerIndex As Integer = 0 To 17
                newRow(innerIndex) = Selected_mail_items(outerIndex, innerIndex)
            Next
            table.Rows.Add(newRow)
        Next

        ext = ".csv"
        Dim FilePath As String = FolderPicker() & "\" & FileName()
        Dim csvTbl As String = ConvertToCSV(table)
        'Dim file As System.IO.StreamWriter

        'file = My.Computer.FileSystem.OpenTextFileWriter(FilePath, True)
        'file.WriteLine(csvTbl)
        'file.Close()
        On Error Resume Next
        IO.File.WriteAllText(FilePath, ConvertToCSV(table))

    End Sub

    Private Function ConvertToCSV(ByVal dt As DataTable) As String
        Dim sb As New Text.StringBuilder()

        For Each row As DataRow In dt.Rows
            sb.AppendLine(String.Join(",", (From i As Object In row.ItemArray Select i.ToString().Replace("""", """""").Replace(",", "\,").Replace(Environment.NewLine, "\" & Environment.NewLine).Replace("\", "\\")).ToArray()))
        Next

        Return sb.ToString()
    End Function

    Public Sub Attachments() 'No Exported Attachments

        Dim myOlApp As Outlook.Application = New Outlook.Application
        Dim objView As Outlook.Explorer = myOlApp.ActiveExplorer
        Dim olMail As Outlook.MailItem
        Dim MailMetadata As Array
        Dim olAttachment As Outlook.Attachment
        Dim i As Integer
        Dim FilePathConverter As String

        filePathPicked = FolderPicker()
        'Set the objView Objext to be the users active Outlook window
        'Common Errors include:
        '   Lack of memory due to a 32 bit system
        '   File not type recognised/corupted
        '   File to large to export due to a 32 bit system
        On Error Resume Next

        'Make a new folder on the users Desktop | Will skip this is it allready exists through the above error handeling
        'MkDir "C:\Users\" & Environ("UserName") & "\Desktop\Attachments"

        'Loop through each selected mail items
        For Each olMail In objView.Selection
            'Loop through each attachment in the selected mail item

            For Each olAttachment In olMail.Attachments
                If olAttachment.FileName <> "" Then
                    'FilePathConverter = File_Exists("C:\Users\" & Environ("UserName") & "\" & FileLocation & "\Attachments\" & FileName & "\" & olAttachment.FileName)
                    FilePathConverter = File_Exists(filePathPicked & "\" & olAttachment.FileName)
                    olAttachment.SaveAsFile(FilePathConverter)
                End If
            Next olAttachment
        Next olMail
    End Sub


    Function FolderPicker() As String
        Dim folderDlg As New FolderBrowserDialog
        Dim str As String = ""
        folderDlg.ShowNewFolderButton = True
        If (folderDlg.ShowDialog() = DialogResult.OK) Then
            str = folderDlg.SelectedPath
            Dim root As Environment.SpecialFolder = folderDlg.RootFolder
        End If
        Return str
    End Function

    Private Sub Mail_Scrape()
        'Scrapes and retrievs all mail items in to a Module level 2D Array
        Call get_Selected_mail_items()
        'Replace all " in the body with ' for file formatting standards
        Call CleanText()
    End Sub

    Private Function FileName() As String
        Dim FileDate As String
        Dim UserName As String
        Dim tempArray As Array
        'Convert the current date to text YYMMDD
        FileDate = Format(Now(), "yymmdd")
        'Convert the users profile name to text
        UserName = Environ("UserName")
        'Split the username by "."
        tempArray = Split(UserName, ".")
        'Initiate teh UserName String variable to be reformed without a "."
        UserName = ""
        'Loop through the User name Array | tempArray()
        For i = 0 To UBound(tempArray)
            'If Last item in array then
            If i = UBound(tempArray) Then
                'Concatenate UserName with the last array item
                UserName = UserName & tempArray(i)
                'Not the last item in the array
            Else
                'Concatenate UserName with the current array item and "_"
                UserName = UserName & tempArray(i) & "_"
            End If
        Next i
        'Retutn fileName by concatenating FileDate-UserName-Mail_Scrape.ext
        FileName = FileDate & "-" & UserName & "-" & "Mail_Scrape" & ext
    End Function

    Private Function jsonArray(str As String, del As String) As String
        Dim tmpArray() As String
        Dim tmpString As String
        'Split up the string
        tmpArray = Split(str, del)
        tmpString = "[" & vbNewLine & vbTab & vbTab & vbTab & vbTab
        For i = 0 To UBound(tmpArray)
            tmpArray(i) = Trim(tmpArray(i))
            If i = UBound(tmpArray) Then
                tmpString = tmpString & "{""email"":""" & tmpArray(i) & """}" & vbNewLine & vbTab & vbTab & vbTab
            Else
                tmpString = tmpString & "{""email"":""" & tmpArray(i) & """}," & vbNewLine & vbTab & vbTab & vbTab & vbTab
            End If
        Next i
        tmpString = tmpString & "]"
        jsonArray = tmpString
    End Function

    Private Function File_Exists(fielPath As String) As String
        Dim strFileExists As String
        Dim fileExists As Boolean
        Dim temp_FileName As String, temp_FileName_Placeholder As String
        Dim temp_FileArray As Array
        Dim temp_FileExt As String
        Dim temp_path As String
        Dim i As Integer
        'Look for the item (filePath)
        strFileExists = Dir(fielPath)
        'Does the file exist
        If strFileExists <> "" Then
            'Breakuup the filepath in to three components
            '   Path | File Name | File Extention
            'Split the filepath into an array by "."
            temp_FileArray = Split(strFileExists, ".")
            'Extract the File Extention by Concatenating "." * the last item in the temp_FileArray()
            temp_FileExt = "." & temp_FileArray(UBound(temp_FileArray))
            'Extract the File Name by through last item in the temp_FileArray()
            temp_FileName = temp_FileArray(0)
            'Resplit the filePath this time by "\"
            temp_FileArray = Split(fielPath, "\")
            'Initilise the temp_path string variable
            temp_path = ""
            'Loop through temp_FileArray() stopping fhort of the last array item
            For i = 0 To UBound(temp_FileArray) - 1
                'Concatenating all teh looped temp_FileArray() items
                temp_path = temp_path & temp_FileArray(i) & "\"
            Next i
            'Initilising the fileExists Boolean Variable which operates as a gate/switch for the below Do While Loop
            fileExists = True
            'Initilise the temp_FileName_Placeholder to be reset and ammended each loop
            temp_FileName_Placeholder = temp_FileName
            'Initilise the counter (i) to be appended to teh file name
            i = 1
            'While fileExists = True rename the variable by concatenating "(" & i & ")"
            Do While fileExists = True
                'Increment the temp_FileName_Placeholder by appending temp_FileName & "(" & i & ")"
                temp_FileName_Placeholder = temp_FileName & "(" & i & ")"
                'Check if teh appended file name exists
                If Dir(temp_path & temp_FileName_Placeholder & temp_FileExt) <> "" Then
                    'Incrument the counter (i)
                    i = i + 1
                Else
                    'Return the new appended fileName
                    fielPath = temp_path & temp_FileName_Placeholder & temp_FileExt
                    'Break teh loop
                    fileExists = False
                End If
            Loop
        Else
            'File does not exist and return fielPath
            fielPath = fielPath
        End If
        'Return teh new or same fiel name
        File_Exists = fielPath
        Return fielPath
    End Function

    Private Sub CleanText()
        Dim i As Single, j As Single
        Dim myString As String
        'Initilise myString as the cleaned string
        myString = ""
        'Loop through all rows (except the header) in the 2D Array | Selected_mail_items()
        For i = 1 To UBound(Selected_mail_items)
            'Loop through each column/dimention in the 2D Array | Selected_mail_items()
            For j = 0 To ArrayDim
                'Replace " with '
                Selected_mail_items(i, j) = Replace(Selected_mail_items(i, j), """", "'")
            Next j
        Next i
    End Sub

    Private Sub get_Selected_mail_items()
        Dim myOlApp As Outlook.Application = New Outlook.Application
        Dim objView As Outlook.Explorer = myOlApp.ActiveExplorer
        Dim oMail As Outlook.MailItem
        Dim i As Integer

        'Initilis the counter i as 1
        i = 1
        'Loop through each selected mail item to get a count to initilise the below 2D array | Selected_mail_items()
        For Each oMail In objView.Selection
            i = i + 1
        Next oMail

        'initilise the 2D Array | Selected_mail_items()
        ReDim Selected_mail_items(0 To i - 1, 0 To ArrayDim)
        'Add headders to the 2D Array | Selected_mail_items(0,?)
        Selected_mail_items(0, 0) = "To"
        Selected_mail_items(0, 1) = "CC"
        Selected_mail_items(0, 2) = "Reply_Recipient_Names"
        Selected_mail_items(0, 3) = "Sender_Email_Address"
        Selected_mail_items(0, 4) = "Sender_Name"
        Selected_mail_items(0, 5) = "Sent_On_Behalf_Of_Name"
        Selected_mail_items(0, 6) = "Sender_Email_Type"
        Selected_mail_items(0, 7) = "Sent"
        Selected_mail_items(0, 8) = "Size"
        Selected_mail_items(0, 9) = "Unread"
        Selected_mail_items(0, 10) = "Creation_Time"
        Selected_mail_items(0, 11) = "Last_Modification_Time"
        Selected_mail_items(0, 12) = "Sent_On"
        Selected_mail_items(0, 13) = "Received_Time"
        Selected_mail_items(0, 14) = "Importance"
        Selected_mail_items(0, 15) = "Received_By_Name"
        Selected_mail_items(0, 16) = "Received_On_Behalf_Of_Name"
        Selected_mail_items(0, 17) = "Subject"
        Selected_mail_items(0, 18) = "Body"
        'Reinitilise that counter (i) to skip the header file
        i = 1
        'Any incompatable mail items are skipped

        'Loop through each selected mail item and add teh metadat to the 2D Array | Selected_mail_items(?>0,?)
        For Each olMail In objView.Selection
            On Error GoTo nxt
            Selected_mail_items(i, 0) = olMail.To
            Selected_mail_items(i, 1) = olMail.CC
            Selected_mail_items(i, 2) = olMail.ReplyRecipientNames
            Selected_mail_items(i, 3) = olMail.SenderEmailAddress
            Selected_mail_items(i, 4) = olMail.SenderName
            Selected_mail_items(i, 5) = olMail.SentOnBehalfOfName
            Selected_mail_items(i, 6) = olMail.SenderEmailType
            Selected_mail_items(i, 7) = olMail.Sent
            Selected_mail_items(i, 8) = olMail.Size
            Selected_mail_items(i, 9) = olMail.UnRead
            Selected_mail_items(i, 10) = olMail.CreationTime
            Selected_mail_items(i, 11) = olMail.LastModificationTime
            Selected_mail_items(i, 12) = olMail.SentOn
            Selected_mail_items(i, 13) = olMail.ReceivedTime
            Selected_mail_items(i, 14) = olMail.Importance
            Selected_mail_items(i, 15) = olMail.ReceivedByName
            Selected_mail_items(i, 16) = olMail.ReceivedOnBehalfOfName
            Selected_mail_items(i, 17) = olMail.Subject
            Selected_mail_items(i, 18) = olMail.Body
            i = i + 1
            'Skipped Mail Item
nxt:
            'Reinitilise error to exit the subroutine is errors persist
            On Error GoTo en
        Next olMail
        'Persistant erros | Exit Sub
en:
        'Reinitilise error handeler to default
        On Error GoTo 0
        'Add Selected_mail_items array to Selected_mail_items (Module Level Array/Variant Variable)

    End Sub

    'READ CSV
    'Using MyReader As New Microsoft.VisualBasic.FileIO.
    'TextFieldParser("c:\" & Environ("Username") & "\desktop\")

    '    MyReader.TextFieldType =
    'Microsoft.VisualBasic.FileIO.FieldType.Delimited
    '    MyReader.Delimiters = New String() {vbTab}
    '    Dim currentRow As String()

    '    currentRow = MyReader.ReadFields()
    'End Using
End Module
