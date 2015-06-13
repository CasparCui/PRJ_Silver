Imports System.Web.Script.Serialization

Namespace SilverCat

    Module SilverCat
        Function Main() As Integer
            Dim baseUri As New Uri("D:\Users\lemac\Source\Repos\PRJ_Silver\SilverCat\SilverCat\")

            Dim exitCode As Integer = -1

            Dim pdfP As PdfProcessing = Nothing
            Try
                pdfP = PdfProcessing.GetInstance
                Dim res() As String

                Dim securityParam As Hashtable
                Dim jsonSecurityParam As String
                '' 受信者の公開鍵で暗号化する場合はこちらのJSONパラメータで暗号化する。
                ''      (New Uri(baseUri, "resource\JsonSecurityParameterRecipient.txt")).AbsolutePath, _
                '' 電子署名用JSONパラメータ
                securityParam = (New JavaScriptSerializer).Deserialize(Of Hashtable)(
                    System.IO.File.ReadAllText( _
                        (New Uri(baseUri, "resource\JsonSecurityParameter.txt")).AbsolutePath, _
                        System.Text.Encoding.UTF8) _
                )
                jsonSecurityParam = (New JavaScriptSerializer).Serialize(securityParam).ToString()

                '' PDFファイル加工指示設定パラメータ
                Dim processingParam As Dictionary(Of String, Dictionary(Of String, Object))
                processingParam = (New JavaScriptSerializer).Deserialize(Of Dictionary(Of String, Dictionary(Of String, Object)))(
                    System.IO.File.ReadAllText( _
                        (New Uri(baseUri, "resource\JsonProcessingParameter.txt")).AbsolutePath, _
                        System.Text.Encoding.UTF8) _
                )

                '' PostScriptファイルからPDFファイル作成
                res = pdfP.CreatePdf((New Uri(baseUri, "data\sample.ps")).AbsolutePath, _
                               (New Uri(baseUri, "data\sample.pdf")).AbsolutePath, _
                               (New Uri(baseUri, "resource\SilverCat.joboptions")).AbsolutePath)
                Console.Error.WriteLine(res(1))

                '' ウォーターマークの設定
                res = pdfP.SetWatermarkPdf(res(0), (New Uri(baseUri, "data\sample_watermark.pdf")).AbsolutePath, processingParam)
                Console.Error.WriteLine(res(1))

                '' 電子署名
                res = pdfP.SignPdf(res(0), (New Uri(baseUri, "data\sample_signed.pdf")).AbsolutePath, jsonSecurityParam)
                Console.Error.WriteLine(res(1))

                '' PostScriptファイルからPDFファイル作成
                res = pdfP.CreatePdf((New Uri(baseUri, "data\sample.ps")).AbsolutePath, _
                               (New Uri(baseUri, "data\sample2.pdf")).AbsolutePath, _
                               (New Uri(baseUri, "resource\SilverCat.joboptions")).AbsolutePath)
                Console.Error.WriteLine(res(1))

                '' ウォーターマークの設定
                res = pdfP.SetWatermarkPdf(res(0), (New Uri(baseUri, "data\sample2_watermark.pdf")).AbsolutePath, processingParam)
                Console.Error.WriteLine(res(1))

                '' セキュリティ設定のみ
                res = pdfP.SetSecurityPdf(res(0), (New Uri(baseUri, "data\sample2_security.pdf")).AbsolutePath, jsonSecurityParam)
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

