Namespace SilverCat

    Module SilverCat
        Private Const BASE_DIR = "D:\Users\lemac\Source\Repos\PRJ_Silver\SilverCat\SilverCat\"

        Function Main() As Integer
            Dim pdfP As PdfProcessing = PdfProcessing.GetInstance

            Dim res() As String
            '' rect(0):x
            '' rect(1):y
            '' rect(2):width
            '' rect(3):height
            Dim sigFieldRect() As Integer = New Integer() {1, 1, 36, 36}

            res = pdfP.CreatePdf(BASE_DIR & "data\sample.ps", _
                           BASE_DIR & "data\sample.pdf", _
                           BASE_DIR & "resource\SilverCat.joboptions")
            Console.Error.WriteLine(res(1))

            res = pdfP.SignPdf(BASE_DIR & "data\sample.pdf", _
                           BASE_DIR & "data\sample_signed.pdf", _
                           "password!", _
                           "D:/Users/lemac/Source/Repos/PRJ_Silver/SilverCat/SilverCat/resource/my_private_key.p12",
                           "PASSWORD_ENCRYPT_POLICY", _
                           sigFieldRect)
            Console.Error.WriteLine(res(1))


            res = pdfP.CreatePdf(BASE_DIR & "data\sample.ps", _
                           BASE_DIR & "data\sample2.pdf", _
                           BASE_DIR & "resource\SilverCat.joboptions")
            Console.Error.WriteLine(res(1))

            res = pdfP.SignPdf(BASE_DIR & "data\sample2.pdf", _
                           BASE_DIR & "data\sample_signed2.pdf", _
                           "password!", _
                           "D:/Users/lemac/Source/Repos/PRJ_Silver/SilverCat/SilverCat/resource/my_private_key.p12",
                           "PASSWORD_ENCRYPT_POLICY", _
                           sigFieldRect)
            Console.Error.WriteLine(res(1))

            Return 0
        End Function
    End Module

End Namespace

