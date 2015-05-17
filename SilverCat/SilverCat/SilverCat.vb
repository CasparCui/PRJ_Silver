Namespace SilverCat

    Module SilverCat
        Private Const BASE_DIR = "D:\Users\lemac\Source\Repos\PRJ_Silver\SilverCat\SilverCat\"

        Function Main() As Integer
            Dim pdfP As PdfProcessing = PdfProcessing.GetInstance


            pdfP.CreatePdf(BASE_DIR & "data\sample.ps", _
                           BASE_DIR & "data\sample.pdf", _
                           BASE_DIR & "resource\SilverCat.joboptions")


            Return 0
        End Function
    End Module

End Namespace

