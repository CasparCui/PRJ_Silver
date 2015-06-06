Imports Acrobat
Imports ACRODISTXLib
Imports AFORMAUTLib
Imports System.IO


Namespace SilverCat

    Public Class PdfProcessing
        Implements IDisposable

        ''' <summary>
        ''' シングルトン。
        ''' </summary>
        Private Shared me_ As PdfProcessing = New PdfProcessing

        ''' <summary>
        ''' Acrobat Distiller。
        ''' </summary>
        Private Shared WithEvents pdfDistiller_ As PdfDistiller

        ''' <summary>
        ''' Acrobat本体。
        ''' </summary>
        Private Shared acroApp_ = New AcroApp

        ''' <summary>
        ''' ログ出力。
        ''' </summary>
        Private Shared log_ As StringWriter

        ''' <summary>
        ''' Acrobat Distillerの処理状態を管理するフラグ。
        ''' </summary>
        ''' <remarks>
        ''' Acrobat Distillerのイベント通知で、処理中ステータスの管理に利用。
        ''' True:処理中。False:アイドル。
        ''' </remarks>
        Private bWorking_ As Boolean

        ''' <summary>
        ''' コンストラクタ(不可視)。
        ''' </summary>
        Private Sub New()
            log_ = New StringWriter

            pdfDistiller_ = New PdfDistiller

            '' Acrobat Distillerのイベントハンドラの設定
            '' PDFが作成されたイベント
            AddHandler pdfDistiller_.OnJobDone, _
                Sub(input As String, output As String)
                    log_.WriteLine(">>Created the PDF[" & output & "].")
                End Sub

            '' ジョブが失敗したイベント
            AddHandler pdfDistiller_.OnJobFail, _
                Sub(input As String, output As String)
                    log_.WriteLine(">>Failed to create the PDF(" & output & ").")
                End Sub

            '' PDF作成中イベント
            AddHandler pdfDistiller_.OnPercentDone, _
                Sub(percentDone As Integer)
                    If percentDone > 0 Then
                        Me.bWorking_ = True
                    End If
                    If Me.bWorking_ Then
                        If percentDone = 0 Then
                            log_.WriteLine("Idle:")
                            Me.bWorking_ = False
                        Else
                            log_.WriteLine(percentDone & "% processing...")
                        End If
                    End If
                End Sub

            '' PDFページ番号イベント
            AddHandler pdfDistiller_.OnPageNumber, _
                Sub(pageNum As Integer)
                    log_.WriteLine("be creating (" & pageNum & ") page...")
                End Sub

            '' True:Acrobat Distillerは画面表示する。
            '' False:Acrobat Distillerは画面表示しない。
            pdfDistiller_.bShowWindow = False

            '' False:FileToPDFはすぐにPDFジョブを処理し、 PDFファイルが作成されるまで戻りません。
            '' True:FileToPDFはDistillerの内部ジョブキューにPDFジョブを送信し、すぐに戻ります。
            '' ジョブは、いくつかの後の時点で処理されます。ジョブが完了したときを知るためには、
            '' Distillerのジョブの処理中に実行されるイベントで確認できます。
            pdfDistiller_.bSpoolJobs = False
        End Sub

        ''' <summary>
        ''' シングルトンを返します。
        ''' </summary>
        Public Shared Function GetInstance() As PdfProcessing
            Return me_
        End Function

#Region "PostScriptファイルからPDFファイルを作成します。"
        ''' <summary>
        ''' PostScriptファイルからPDFファイルを作成します。
        ''' </summary>
        ''' <param name="inPsFilePath">入力PostScriptファイルパス。</param>
        ''' <param name="outPdfFilePath">出力PDFファイルパス。</param>
        ''' <param name="jobOptionFilePath">Acrobat Distiller ジョブオプションファイルパス。</param>
        ''' <returns>
        ''' String(0):PDFファイルのフルパス。
        ''' String(1):Acrobat Distiller ログメッセージ文字列。
        ''' </returns>
        ''' <remarks>
        ''' 参考URL：
        ''' http://help.adobe.com/livedocs/acrobat_sdk/9.1/Acrobat9_1_HTMLHelp
        ''' </remarks>
        Public Function CreatePdf(ByVal inPsFilePath As String, _
                             ByVal outPdfFilePath As String, _
                             ByVal jobOptionFilePath As String) As String()

            Dim result() As String = New String() {Nothing, Nothing}
            Try
                log_.Write(">>Input PostScript file is (")
                log_.Write(inPsFilePath)
                log_.WriteLine(").")

                '' PDF作成中イベントハンドラ用に処理中フラグをFalseにして、
                '' まずは処理は止まっていることにする。
                Me.bWorking_ = False

                '' PDFファイルをファイル
                Dim rc As Integer
                rc = pdfDistiller_.FileToPDF(inPsFilePath, outPdfFilePath, jobOptionFilePath)
                Select Case rc
                    Case -1
                        log_.WriteLine("An exception occurred in the [PdfDistiller.bSpoolJobs].")
                        Throw New ArgumentException(log_.ToString())
                    Case 0
                        log_.WriteLine("An exception occurred in the [PdfDistiller.FileToPDF()] method. Please check the argument.")
                        Throw New ArgumentException(log_.ToString())
                    Case 1
                        '' OK
                        log_.WriteLine("Successful in the creation of [" & outPdfFilePath & "] pdf file.")
                    Case Else
                        log_.WriteLine("It failed to execute [PdfDistiller.FileToPDF()] method.")
                        Throw New Exception(log_.ToString())
                End Select


                result(0) = outPdfFilePath
                result(1) = log_.ToString()
            Finally
            End Try
            Return result
        End Function
#End Region

#Region "PDFファイルに電子署名します。"
        ''' <summary>
        ''' PDFファイルに電子署名します。
        ''' </summary>
        ''' <paramref name="inPdfFilePath">入力PDFファイルへのフルパス。</paramref>
        ''' <paramref name="outPdfFilePath">出力PDFファイルへのフルパス。</paramref>
        ''' <paramref name="jsonSignParam">電子署名Javascriptに渡す引数。JSON形式のString。</paramref>
        ''' JavaScriptで電子署名をするので、電子証明書ファイルパスはjavascriptで解釈できるファイルパスの区切り文字"/"のこと。</param>
        ''' <remarks>
        ''' http://kb2.adobe.com/jp/cps/511/511344/attachments/511344_Security_2005_DevCon.pdf
        ''' %Acrobatインストールディレクトリ%Acrobat\Javascriptsに、SignPdf.jsを事前に配置のこと。
        ''' </remarks>
        ''' <returns>
        ''' String(0):出力PDFファイルのフルパス。
        ''' String(1):ログメッセージ文字列。
        ''' </returns>
        Public Function SignPdf(ByVal inPdfFilePath As String, _
                                ByVal outPdfFilePath As String, _
                                ByVal jsonSignParam As String) As String()

            Dim inPdDoc As AcroPDDoc = Nothing
            Dim inAvDoc As AcroAVDoc = Nothing
            Dim formApp As AFormApp = Nothing
            Dim fields As Fields = Nothing
            Dim rc As Boolean = False

            Dim result() As String = New String() {Nothing, Nothing}
            Try
                log_.Write(">>Input Pdf file is (")
                log_.Write(inPdfFilePath)
                log_.WriteLine(")")



                '' 電子署名する前の入力PDFファイルのまま残しておきたいので、コピーする。
                File.Copy(inPdfFilePath, outPdfFilePath, True)

                inPdDoc = New AcroPDDoc()

                rc = inPdDoc.Open(outPdfFilePath)

                If Not (rc) Then
                    Throw New IOException(">>It failed to open, for sign file(" & outPdfFilePath & ") ")
                End If

                inAvDoc = CType(inPdDoc.OpenAVDoc("TEMP_PDF_DOCUMENT"), AcroAVDoc)

                '' 電子署名自体はJavascriptで行うので、ここではJavascriptを呼び出すイメージ。
                formApp = New AFormApp()

                fields = CType(formApp.Fields, Fields)

                Dim nVersion As String
                nVersion = fields.ExecuteThisJavascript("event.value = app.viewerVersion;")
                log_.WriteLine("The Acrobat viewer version is " & nVersion & ".")


                Dim jsCode As New StringWriter
                jsCode.Write("SignPdf(this, ")
                jsCode.Write(jsonSignParam)
                jsCode.Write(");")

                Dim jsRc As String = fields.ExecuteThisJavascript(jsCode.ToString())

                '' 戻り値:ゼロが正常終了。それ以外は、異常終了。
                If (jsRc <> "0") Then
                    Throw New Exception("SignPdf:" & jsRc)
                End If

                '' AVDoc.Close(bNoSave)
                '' bNoSave:
                '' 正の数:PDFドキュメントを保存しないで閉じます。 
                '' 0でPDFドキュメントが変更されていた:保存するかどうか確認するダイアログを表示
                '' 負の数:確認なしにPDFドキュメントを保存。
                rc = inAvDoc.Close(-1)

                If rc Then
                    log_.WriteLine(">>digital sign pdf document to save (" & outPdfFilePath & ").")
                Else
                    Throw New IOException("It failed digital sign pdf document to save (" + outPdfFilePath + ").")
                End If

                result(0) = outPdfFilePath
                result(1) = log_.ToString()

            Catch ex As Exception
                '' 電子署名用に作ったPDFドキュメントを保存せずに閉じます。
                rc = inAvDoc.Close(1)
                '' 仮に保存したファイルを削除します。
                File.Delete(outPdfFilePath)
                Throw ex
            Finally
                If Not (fields Is Nothing) Then
                    Marshal.ReleaseComObject(fields)
                    fields = Nothing
                End If
                If Not (formApp Is Nothing) Then
                    Marshal.ReleaseComObject(formApp)
                    formApp = Nothing
                End If
                If Not (inAvDoc Is Nothing) Then
                    inAvDoc.Close(1)
                    Marshal.ReleaseComObject(inAvDoc)
                    inAvDoc = Nothing
                End If
                If Not (inPdDoc Is Nothing) Then
                    inPdDoc.Close()
                    Marshal.ReleaseComObject(inPdDoc)
                    inPdDoc = Nothing
                End If
            End Try

            Return result

        End Function
#End Region


#Region "Dispose"
        ''' <summary>
        ''' リソース解放
        ''' </summary>
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            Console.Error.WriteLine(">PdfProcessing:Resources of the release.")
            If Not (pdfDistiller_ Is Nothing) Then
                Marshal.ReleaseComObject(pdfDistiller_)
                pdfDistiller_ = Nothing
            End If
            If Not (acroApp_ Is Nothing) Then
                acroApp_.CloseAllDocs()
                acroApp_.Exit()
                Marshal.ReleaseComObject(acroApp_)
                acroApp_ = Nothing
            End If
        End Sub
#End Region
#Region "IDisposable Support "
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class


End Namespace




