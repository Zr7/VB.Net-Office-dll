Imports System
Imports System.IO
'De Itextsharp dll
Imports iTextSharp.text.xml
Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports iTextSharp
Imports System.Windows.forms
'Deze Class is nog in ontwikkeling en kan nog niet echt gebruikt worden

Public Class PdfClass
    Sub PrintPDF(ByVal PdFileName As String)
        Dim psiPrint As New System.Diagnostics.ProcessStartInfo()
        psiPrint.Verb = "print"
        psiPrint.WindowStyle = ProcessWindowStyle.Hidden
        psiPrint.FileName = PdFileName
        psiPrint.UseShellExecute = True
        System.Diagnostics.Process.Start(psiPrint)
    End Sub

    Function GetFields(ByVal PdFileName As String)
        Dim pdfTemplate As String = PdFileName
        Try
            Dim readerPDF As New PdfReader(pdfTemplate)

            Dim PDFfld As Object
            Using MyBox As New Windows.Forms.Button
                For Each PDFfld In readerPDF.AcroFields.Fields
                    Return PDFfld.key.ToString()
                Next

            End Using


        Catch ex As IOException
            MsgBox(ex.ToString)
        End Try
    End Function
    'Now when you select the pdf file and click on the “Get Form Fields” button, you will notice the textbox populates with the Form Field names.




End Class
