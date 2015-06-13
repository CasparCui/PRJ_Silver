Imports System.Web.Script.Serialization

Namespace SilverCat

    Module SilverCat
        Private Const BASE_DIR = "D:\Users\lemac\Source\Repos\PRJ_Silver\SilverCat\SilverCat\"

        Function Main() As Integer

            Dim baseUri As New Uri("D:\Users\lemac\Source\Repos\PRJ_Silver\SilverCat\SilverCat\")

            Dim exitCode As Integer = -1

            Dim pdfP As PdfProcessing = Nothing
            Try
                pdfP = PdfProcessing.GetInstance
                Dim res() As String
                Dim securityParam As Hashtable = New Hashtable
                Dim jsonSecurityParam As String
                securityParam.Add("x", 1)
                securityParam.Add("y", 1)
                securityParam.Add("width", 36)
                securityParam.Add("height", 36)
                securityParam.Add("digitalIdFilePath", (New Uri(baseUri, "resource\my_private_key.p12")).AbsolutePath)
                securityParam.Add("password", "password!")
                '' AcrobatPro「表示(V)」「ツール(T)」「保護(R)」「暗号化」「セキュリティポリシーを管理(M)...」に
                '' "PASSWORD_ENCRYPT_POLICY"を事前に設定のこと。この文字列でjavascript内でセキュリティ設定をします。
                securityParam.Add("securityPolicyName", "PASSWORD_ENCRYPT_POLICY")
                '' AcrobatPro「編集(E)」「環境設定(N)...」「分類(G)」「署名」
                '' 「電子署名」「作成と表示方法」「詳細...」「表示方法」に
                '' "AppearanceSilverCatSignature"を事前に設定のこと。この文字列でjavascript内で
                '' 電子署名の表示方法を設定するため。
                securityParam.Add("appearanceSignature", "AppearanceSilverCatSignature")

                'Dim recipientPublicCert As ArrayList = New ArrayList
                'recipientPublicCert.Add((New Uri(baseUri, "resource\CertExchangeHanako.cer")).AbsolutePath)
                'recipientPublicCert.Add((New Uri(baseUri, "resource\CertExchangeTaro.cer")).AbsolutePath)
                'securityParam.Add("recipientPublicCerts", recipientPublicCert)

                '' 受信者公開鍵で暗号化せず、セキュリティポリシーで暗号化するので、Nothing
                securityParam.Add("recipientPublicCerts", Nothing)

                jsonSecurityParam = (New JavaScriptSerializer).Serialize(securityParam).ToString()

                res = pdfP.CreatePdf((New Uri(baseUri, "data\sample.ps")).AbsolutePath, _
                               (New Uri(baseUri, "data\sample.pdf")).AbsolutePath, _
                               (New Uri(baseUri, "resource\SilverCat.joboptions")).AbsolutePath)
                Console.Error.WriteLine(res(1))

                res = pdfP.SignPdf(res(0), (New Uri(baseUri, "data\sample_signed.pdf")).AbsolutePath, jsonSecurityParam)
                Console.Error.WriteLine(res(1))


                res = pdfP.CreatePdf((New Uri(baseUri, "data\sample.ps")).AbsolutePath, _
                               (New Uri(baseUri, "data\sample2.pdf")).AbsolutePath, _
                               (New Uri(baseUri, "resource\SilverCat.joboptions")).AbsolutePath)
                Console.Error.WriteLine(res(1))
                '' sample2はセキュリティ設定のみ
                res = pdfP.SetSecurityPdf(res(0), (New Uri(baseUri, "data\sample_signed2.pdf")).AbsolutePath, jsonSecurityParam)
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

