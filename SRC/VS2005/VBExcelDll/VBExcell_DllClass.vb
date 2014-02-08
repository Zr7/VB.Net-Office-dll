'
'Dll class om het gebruik van excell en itextSharp te vergemakkelijken , 
'waardoor er minimale referenties nodig zijn in andere programma's
'
'Deze Dll bevat ook de benodigde excell bestanden die nodig zijn voor 
'bepaalde programma's 
'het eerste deel basic bevat bestand bewerking , 
'het tweede deel excel bewerking ,
' het derde deel email
'hetvierde deel is gereserveerd voor pdf bewerking
'het (voorlopig in ontwikkeling ) deel bevat Database bewerking
'Alle code is nog in bewerking
'
'' CreateExcell is vervallen
'
' ToDo  Finalise iText
Option Strict Off
Option Explicit Off
'declaratie van de te gebruiken dll's
Imports Microsoft.Office.Core
Imports excel = Microsoft.Office.Interop.Excel

Imports Pdf = PdfClass
'De dll voor printen
Imports System.Drawing
Imports System.Drawing.Printing
'de dll's voor error en bestands behandeling
Imports System
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Diagnostics
Imports System.Reflection
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Globalization
Imports System.Collections
Imports System.Windows.Forms

'De "RootNamespce van de applicatie is bewust leeg gelaten en wordt pas
'hier gedeclareerd , dit om ingenestlede of Dubbele Namespace te voorkomen
Namespace VBExcell
    'Basis bestands behandeling
    Public Class Basic
        Public Sub RetrieveFile(ByVal FileName As String, ByVal ByteName As String)
            '
            'RetrieveFile haalt bestand uit de resourcefile van DIT project , dus dll bestand
            'Project -> Project Properties -> resources-> -> Files -> add file -> FileName *.*
            '
            'Dit is een voorbeeld om resources te gebruiken
            '
            Select Case (ByteName) 'Gebruik "Select Case" om het juiste bestand te selecteren
                Case "Vopak"
                    'WriteAllbytes is voor het ophalen van niet grafische bestanden
                    IO.File.WriteAllBytes(FileName, My.Resources.Vopak)
                Case "Reisplanning" 'Haal een Blanco ReisPlanning formulier op
                    IO.File.WriteAllBytes(FileName, My.Resources.Reisplanning)
                    '
                    'Deze lijst kan dus in theorie oneindig worden
                    'houdt er wel rekening mee , dat dan ook je Dll bestand Groter wordt !!
                    '
                    '
                    'Het is aan te raden om dan niet elke Dll met dezelfde NameSpace te beginnen
                    'omdat je dan naams conflicten krijgt bij declaratie en ophalen
                    '
            End Select
        End Sub
        Function DeleteFile(ByVal FileNm As String) As Boolean
            If (IO.File.Exists(FileNm)) Then
                IO.File.Delete(FileNm)
            End If
        End Function

        Function App_Path() As String
            'Geef de Foldernaam weer van de applicatie
            Return System.AppDomain.CurrentDomain.BaseDirectory()
        End Function
        Function FileExists(ByVal FileFullPath As String) _
                     As Boolean
            'Kijk of het bestand bestaat
            Dim f As New IO.FileInfo(FileFullPath)
            Return f.Exists

        End Function

        Public Function FolderExists(ByVal FolderPath As String) _
       As Boolean

            Dim f As New IO.DirectoryInfo(FolderPath)
            Return f.Exists

        End Function
        Public Sub CreateFolder(ByVal FolderPath As String)
            IO.Directory.CreateDirectory(FolderPath)
        End Sub

    End Class

    Public Class Bexcell
        Inherits Basic 'neem basis declaraties over
        Dim oldCI As System.Globalization.CultureInfo = _
                             System.Threading.Thread.CurrentThread.CurrentCulture()
        Dim xlApp As New excel.Application
        Dim xlWorkBook As excel.Workbook = Nothing
        Dim xlWorkSheet As excel.Worksheet = Nothing
        Dim NoError As Boolean
        Dim CurPath As String
        Public Sub OpenExcell(ByVal NameBook As String)
            'Open excel, en controleer de versie
            On Error Resume Next
            If (Not (Path.GetFullPath(NameBook)) = Nothing) Then
                If (File.Exists(NameBook)) Then
                    On Error GoTo English
                    xlWorkBook = xlApp.Workbooks.Open(Path.GetFullPath(NameBook), , False)
                    NoError = True
                    Resume GoNext 'als er geen error is, ga dan verder met huidige excel
English:
                    On Error Resume Next
                    'Dit stukje zorgt ervoor dat engelse versie van Office/excel gebruikt kan worden
                    'Want niet elke gebruiker heeft een Nederlandse Versie
                    System.Threading.Thread.CurrentThread.CurrentCulture = _
                        New System.Globalization.CultureInfo("en-US")
                    'Verander systeem land naar engels excel-versie
                    xlWorkBook = xlApp.Workbooks.Open(Path.GetFullPath(NameBook), , False)
                    NoError = False 'er is een error geweest, anders werkt deze code niet
                    Resume GoNext
GoNext:
                    On Error Resume Next
                Else
                    'laat een waarschuwing zien als het bestand niet bestaat
                    MessageBox.Show("File does not exist.", "No File", MessageBoxButtons.OK, _
                    MessageBoxIcon.Information)
                    CloseExcell()
                End If
            End If
        End Sub

        Public Sub EditCells(ByVal NameSheet As String, ByVal Collumname As Integer, _
                                    ByVal Rowname As Integer, ByVal EditText As String, ByVal Wrap As Boolean)
            'Excell Blad bewerking met x,y Cel referentie
            xlWorkSheet = xlApp.Worksheets(NameSheet)
            xlWorkSheet.Cells.WrapText = Wrap
            'bewerk de cell met de nieuwe waarde
            xlWorkSheet.Cells(Rowname, Collumname) = EditText

        End Sub

        Sub EditCelRange(ByVal NmSheet As String, ByVal Rng As String, _
                                                 ByVal EdText As String)
            'Excell Blad bewerking met A1 Cel referentie
            xlWorkSheet = xlApp.Worksheets(NmSheet)
            xlWorkSheet.Range(Rng).Value = EdText

        End Sub

        Public Sub PrintExcell(ByVal PrSheet As String)
            ' 
            'Deze Routine werkt TOTAAL onAfhankelijk van de Form
            'Dus heeft geen PrintDialog toevoeging op de Form Nodig
            '
            Dim prSelect As New PrintDialog
            Dim printDoc As New PrintDocument() 'Object referentie voor een leeg document
            Dim Selected_Printer As String
            Dim x As Integer
            xlWorkSheet = xlApp.Worksheets(PrSheet)

            printDoc.DocumentName = xlWorkSheet.Name
            prSelect.Document = printDoc

            prSelect.AllowSelection = True
            prSelect.AllowCurrentPage = True
            prSelect.AllowSomePages = True

            'De eerder gedeclareerde PrintDialog MOET via UseEXDialog op het scherm 
            'weer gegeven worden , anders wacht het op een gebruikers invoering
            ' van een onbereikbaar dialoog scherm
            '
            prSelect.UseEXDialog = True
            'Gebruik Try Catch , want de print opdracht kan natuurlijk ook gecanseld 
            'of worden afgesloten zonder te printen
            Try
                If (prSelect.ShowDialog() = DialogResult.OK) Then
                    'Het x aantal kopiën selecteren via dialoog
                    x = prSelect.PrinterSettings.Copies
                    'Er wordt via de print-dialog een printer gekozen
                    Selected_Printer = prSelect.PrinterSettings.PrinterName
                    xlWorkSheet.PrintOutEx(, , x, , Selected_Printer)
                End If
            Catch Ex As Exception

            Finally
                ReleaseObject(prSelect)
                ReleaseObject(printDoc)
                Selected_Printer = Nothing
            End Try
        End Sub

        Sub DirectPrintExcell(ByVal PrSheet As String, ByVal x As Integer)
            ' 
            'Deze Routine werkt TOTAAL onAfhankelijk van de Form
            'Dus heeft geen PrintDialog toevoeging op de Form Nodig
            '
            Dim prSelect As New PrintDialog
            Using printDoc As New PrintDocument() 'Object referentie voor een leeg document
                xlWorkSheet = xlApp.Worksheets(PrSheet)
                'En we geven de opdracht om het lege document te vullen met het excell blad
                printDoc.DocumentName = xlWorkSheet.Name
                prSelect.Document = printDoc
                Try
                    xlWorkSheet.PrintOutEx(, , x, , )
                Catch Ex As Exception

                Finally
                    ReleaseObject(prSelect)
                End Try
            End Using
        End Sub
        Function GetCellRangeValue(ByVal Nmsheet As String, ByVal Rng As String) As String
            'geeft de waardes weer van een cel van een excel blad
            On Error Resume Next
            xlWorkSheet = xlApp.Worksheets(Nmsheet)
            Return xlWorkSheet.Range(Rng).Value2
        End Function

        Function GetCellValue(ByVal NmSheet As String, ByVal Rnm As Integer, _
                                    ByVal Cnm As Integer) As String
            'geeft de waardes weer van een cel van een excel blad
            Dim Val As String
            On Error Resume Next
            xlWorkSheet = xlApp.Worksheets(NmSheet)
            'Val. = xlWorkSheet.Cells(Rnm, Cnm)
            Return xlWorkSheet.Cells(Rnm, Cnm).Value

        End Function

        Sub CloseExcell()
            'sluit excell en bewaar de veranderingen niet
            'de systeem landinstellingen worden terug gezet
            On Error Resume Next
            xlWorkBook.Close(False)
            xlApp.Quit()

            If NoError = False Then
                System.Threading.Thread.CurrentThread.CurrentCulture = oldCI  'Zet de landinstellingen terug
                'Opm. de systeem landinstellingen worden terug gezet, maar pas NA het afsluiten van excel
            End If


            'ReleaseObject(xlApp)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlWorkSheet)
            ReleaseObject(NoError)
        End Sub
        Public Sub CloseSaveAs(ByVal FileNm As String)
            'sluit excell , bewaar de veranderingen en sla het werkboek op onder een andere naam
            Try
                xlWorkBook.Close(True, FileNm, )
                xlApp.Quit()
                If NoError = False Then 'de systeem landinstellingen worden terug gezet
                    System.Threading.Thread.CurrentThread.CurrentCulture = oldCI  'Zet de landinstellingen terug
                    'Opm. de systeem landinstellingen worden terug gezet, maar pas NA het afsluiten van excel
                End If
            Catch ex As System.Exception
                MessageBox.Show(ex.Message & "=======  Fout Tijdens opslaan:  ======")
            Finally
                'ReleaseObject(xlApp)
                ReleaseObject(xlWorkBook)
                ReleaseObject(xlWorkSheet)
                ReleaseObject(NoError)
            End Try
        End Sub
        Sub CloseSavePdf(ByVal Nmbook As String)

            Dim misValue As Object = System.Reflection.Missing.Value

            Try
                'xlWorkBook = xlApp.Workbooks.Open(Nmbook)
                xlWorkBook.ExportAsFixedFormat(excel.XlFixedFormatType.xlTypePDF _
                , Nmbook, excel.XlFixedFormatQuality.xlQualityStandard, _
                misValue, misValue, misValue, misValue, True)
                'excell bestand NIET opslaan , maar gewoon afsluiten
                CloseExcell()

            Catch ex As System.IO.IOException
                MessageBox.Show("\=======  WRITE TO PDF ERROR:  ======" _
                & Chr(10) & "Er wordt geen excell Bestand bewaard")
                CloseExcell()
            Catch ex As System.Exception
                MessageBox.Show("\n\n=======  WRITE TO PDF ERROR:  ======\n\n" _
                & Chr(10) & "Er wordt geen excell Bestand bewaard")
                '
                'Ondanks de fout , niet vergeten excell af te sluiten 
                '
                CloseExcell()

            Finally
                ' geheugen vrijgeven
                ReleaseObject(misValue)
            End Try
        End Sub
        Public Sub ReleaseObject(ByVal obj As Object)
            'gebruikt geheugen vrij maken

            Try
                While (System.Runtime.InteropServices.Marshal.ReleaseComObject(obj) > 0)
                    'obj = Nothing
                End While
            Catch ex As Exception
                obj = Nothing
            Finally
                obj = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub
    End Class

    Namespace SendFile
        Public Class MAPI
            'Deze Class behandeld en verstuurt emails vanuit een programma
            'Niet over na denken , gewoon gebruiken
            Public Function AddRecipientTo(ByVal email As String) As Boolean
                Return AddRecipient(email, howTo.MAPI_TO)
            End Function

            Public Function AddRecipientCC(ByVal email As String) As Boolean
                Return AddRecipient(email, howTo.MAPI_TO)
            End Function

            Public Function AddRecipientBCC(ByVal email As String) As Boolean
                Return AddRecipient(email, howTo.MAPI_TO)
            End Function

            Public Sub AddAttachment(ByVal strAttachmentFileName As String)
                m_attachments.Add(strAttachmentFileName)
            End Sub

            Public Function SendMailPopup(ByVal strSubject As String, _
                    ByVal strBody As String) As Integer
                Return SendMail(strSubject, strBody, MAPI_LOGON_UI Or MAPI_DIALOG)
            End Function

            Public Function SendMailDirect(ByVal strSubject As String, _
                ByVal strBody As String) As Integer
                Return SendMail(strSubject, strBody, MAPI_LOGON_UI)
            End Function


            <DllImport("MAPI32.DLL")> _
            Private Shared Function MAPISendMail(ByVal sess As IntPtr, _
                ByVal hwnd As IntPtr, ByVal message As MapiMessage, _
                ByVal flg As Integer, ByVal rsv As Integer) As Integer
            End Function

            Private Function SendMail(ByVal strSubject As String, _
                ByVal strBody As String, ByVal how As Integer) As Integer
                Dim msg As MapiMessage = New MapiMessage()
                msg.subject = strSubject
                msg.noteText = strBody

                msg.recips = GetRecipients(msg.recipCount)
                msg.files = GetAttachments(msg.fileCount)

                m_lastError = MAPISendMail(New IntPtr(0), New IntPtr(0), msg, how, 0)
                If m_lastError > 1 Then
                    MessageBox.Show("MAPISendMail failed! " + GetLastError(), _
                        "MAPISendMail")
                End If

                Cleanup(msg)
                Return m_lastError
            End Function

            Private Function AddRecipient(ByVal email As String, _
                ByVal howTo As howTo) As Boolean
                Dim recipient As MapiRecipDesc = New MapiRecipDesc()
                'Voeg een email ontvanger toe
                recipient.recipClass = CType(howTo, Integer)
                recipient.name = email
                m_recipients.Add(recipient)

                Return True
            End Function

            Private Function GetRecipients(ByRef recipCount As Integer) As IntPtr
                recipCount = 0
                If m_recipients.Count = 0 Then
                    Return 0
                End If

                Dim size As Integer = Marshal.SizeOf(GetType(MapiRecipDesc))
                Dim intPtr As IntPtr = Marshal.AllocHGlobal( _
                    m_recipients.Count * size)

                Dim ptr As Integer = CType(intPtr, Integer)
                Dim mapiDesc As MapiRecipDesc
                For Each mapiDesc In m_recipients
                    Marshal.StructureToPtr(mapiDesc, CType(ptr, IntPtr), False)
                    ptr += size
                Next

                recipCount = m_recipients.Count
                Return intPtr
            End Function

            Private Function GetAttachments(ByRef fileCount As Integer) As IntPtr
                fileCount = 0
                If m_attachments Is Nothing Then
                    'bepaal of er een bestand is
                    Return 0
                End If

                If (m_attachments.Count <= 0) Or (m_attachments.Count > _
                    maxAttachments) Then
                    Return 0
                End If
                'De grote van het toegoegde bestand bepalen 
                Dim size As Integer = Marshal.SizeOf(GetType(MapiFileDesc))
                Dim intPtr As IntPtr = Marshal.AllocHGlobal( _
                    m_attachments.Count * size)

                Dim mapiFileDesc As MapiFileDesc = New MapiFileDesc()
                mapiFileDesc.position = -1
                Dim ptr As Integer = CType(intPtr, Integer)
                'Geselecteerd bestand Toevoegen aan de mail 
                Dim strAttachment As String
                For Each strAttachment In m_attachments
                    mapiFileDesc.name = Path.GetFileName(strAttachment)
                    mapiFileDesc.path = strAttachment
                    Marshal.StructureToPtr(mapiFileDesc, CType(ptr, IntPtr), False)
                    ptr += size
                Next

                fileCount = m_attachments.Count
                Return intPtr
            End Function

            Private Sub Cleanup(ByRef msg As MapiMessage)
                Dim size As Integer = Marshal.SizeOf(GetType(MapiRecipDesc))
                Dim ptr As Integer = 0

                If msg.recips <> IntPtr.Zero Then
                    ptr = CType(msg.recips, Integer)
                    Dim i As Integer
                    For i = 0 To msg.recipCount - 1 Step i + 1
                        Marshal.DestroyStructure(CType(ptr, IntPtr), _
                            GetType(MapiRecipDesc))
                        ptr += size
                    Next
                    Marshal.FreeHGlobal(msg.recips)
                End If

                If msg.files <> IntPtr.Zero Then
                    size = Marshal.SizeOf(GetType(MapiFileDesc))

                    ptr = CType(msg.files, Integer)
                    Dim i As Integer
                    For i = 0 To msg.fileCount - 1 Step i + 1
                        Marshal.DestroyStructure(CType(ptr, IntPtr), _
                            GetType(MapiFileDesc))
                        ptr += size
                    Next
                    Marshal.FreeHGlobal(msg.files)
                End If

                m_recipients.Clear()
                m_attachments.Clear()
                m_lastError = 0
            End Sub

            Public Function GetLastError() As String
                If m_lastError <= 26 Then
                    Return errors(m_lastError)
                End If
                Return "MAPI error [" + m_lastError.ToString() + "]"
            End Function

            ReadOnly errors() As String = New String() {"OK [0]", "User abort [1]", _
                "General MAPI failure [2]", "MAPI login failure [3]", _
                "Disk full [4]", "Insufficient memory [5]", "Access denied [6]", _
                "-unknown- [7]", "Too many sessions [8]", _
                "Too many files were specified [9]", _
                "Too many recipients were specified [10]", _
                "A specified attachment was not found [11]", _
                "Attachment open failure [12]", "Attachment write failure [13]", _
                "Unknown recipient [14]", "Bad recipient type [15]", _
                "No messages [16]", "Invalid message [17]", "Text too large [18]", _
                "Invalid session [19]", "Type not supported [20]", _
                "A recipient was specified ambiguously [21]", _
                "Message in use [22]", "Network failure [23]", _
                "Invalid edit fields [24]", "Invalid recipients [25]", _
                "Not supported [26]"}

            Dim m_recipients As New List(Of MapiRecipDesc)
            Dim m_attachments As New List(Of String)
            Dim m_lastError As Integer = 0

            Private Const MAPI_LOGON_UI As Integer = &H1
            Private Const MAPI_DIALOG As Integer = &H8
            Private Const maxAttachments As Integer = 20

            Enum howTo
                MAPI_ORIG = 0
                MAPI_TO
                MAPI_CC
                MAPI_BCC
            End Enum

        End Class

        <StructLayout(LayoutKind.Sequential)> _
        Public Class MapiMessage
            Public reserved As Integer
            Public subject As String
            Public noteText As String
            Public messageType As String
            Public dateReceived As String
            Public conversationID As String
            Public flags As Integer
            Public originator As IntPtr
            Public recipCount As Integer
            Public recips As IntPtr
            Public fileCount As Integer
            Public files As IntPtr
        End Class

        <StructLayout(LayoutKind.Sequential)> _
        Public Class MapiFileDesc
            Public reserved As Integer
            Public flags As Integer
            Public position As Integer
            Public path As String
            Public name As String
            Public type As IntPtr
        End Class

        <StructLayout(LayoutKind.Sequential)> _
        Public Class MapiRecipDesc
            Public reserved As Integer
            Public recipClass As Integer
            Public name As String
            Public address As String
            Public eIDSize As Integer
            Public enTryID As IntPtr
        End Class
        'Einde Email Class / NameSpace
    End Namespace


End Namespace
