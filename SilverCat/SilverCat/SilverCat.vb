Imports System.Web.Script.Serialization

Namespace SilverCat

    Module SilverCat
        Private Const BASE_DIR = "D:\Users\lemac\Source\Repos\PRJ_Silver\SilverCat\SilverCat\"

        Function Main() As Integer
            Dim exitCode As Integer = -1

            Dim pdfP As PdfProcessing = Nothing
            Try
                pdfP = PdfProcessing.GetInstance
                Dim res() As String
                Dim signParam As Hashtable = New Hashtable
                Dim jsonSignParam As String
                signParam.Add("x", 1)
                signParam.Add("y", 1)
                signParam.Add("width", 36)
                signParam.Add("height", 36)
                signParam.Add("digitalIdFilePath", "D:/Users/lemac/Source/Repos/PRJ_Silver/SilverCat/SilverCat/resource/my_private_key.p12")
                signParam.Add("password", "password!")
                signParam.Add("securityPolicyName", "PASSWORD_ENCRYPT_POLICY")

                Dim recipientPublicCert As ArrayList = New ArrayList
                recipientPublicCert.Add("D:/Users/lemac/Source/Repos/PRJ_Silver/SilverCat/SilverCat/resource/CertExchangeHanako.cer")
                recipientPublicCert.Add("D:/Users/lemac/Source/Repos/PRJ_Silver/SilverCat/SilverCat/resource/CertExchangeTaro.cer")
                signParam.Add("recipientPublicCerts", recipientPublicCert)
                jsonSignParam = (New JavaScriptSerializer).Serialize(signParam).ToString()

                res = pdfP.CreatePdf(BASE_DIR & "data\sample.ps", _
                               BASE_DIR & "data\sample.pdf", _
                               BASE_DIR & "resource\SilverCat.joboptions")
                Console.Error.WriteLine(res(1))

                res = pdfP.SignPdf(BASE_DIR & "data\sample.pdf", BASE_DIR & "data\sample_signed.pdf", jsonSignParam)
                Console.Error.WriteLine(res(1))


                res = pdfP.CreatePdf(BASE_DIR & "data\sample.ps", _
                               BASE_DIR & "data\sample2.pdf", _
                               BASE_DIR & "resource\SilverCat.joboptions")
                Console.Error.WriteLine(res(1))

                res = pdfP.SignPdf(BASE_DIR & "data\sample2.pdf", BASE_DIR & "data\sample_signed2.pdf", jsonSignParam)
                Console.Error.WriteLine(res(1))

                exitCode = 0
            Catch ex As Exception
                Debug.Print(ex.ToString())
            Finally
                If IsNothing(pdfP) Then
                    '' nop
                Else
                    pdfP.Dispose()
                End If

            End Try

            Return exitCode
        End Function
    End Module

End Namespace

