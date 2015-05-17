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

        ''' <summary>
        ''' PostScriptファイルからPDFファイルを作成します。
        ''' </summary>
        ''' <param name="inPS"></param>
        ''' <param name="outPDF"></param>
        ''' <param name="jobOptionFilePath"></param>
        ''' <returns>
        ''' String(0):PDFファイルのフルパス。
        ''' String(1):Acrobat Distiller ログメッセージ文字列。
        ''' </returns>
        ''' <remarks>
        ''' 参考URL：
        ''' http://help.adobe.com/livedocs/acrobat_sdk/9.1/Acrobat9_1_HTMLHelp
        ''' </remarks>
        Public Function CreatePdf(ByRef inPS As String, _
                             ByRef outPDF As String, _
                             ByRef jobOptionFilePath As String) As String()

            Dim result() As String = New String() {Nothing, Nothing}
            Try
                log_.Write(">>Input PostScript file is (")
                log_.Write(inPS)
                log_.WriteLine(").")

                '' PDF作成中イベントハンドラ用に処理中フラグをFalseにして、
                '' まずは処理は止まっていることにする。
                Me.bWorking_ = False

                '' PDFファイルをファイル
                Dim rc As Integer
                rc = pdfDistiller_.FileToPDF(inPS, outPDF, jobOptionFilePath)
                Select Case rc
                    Case -1
                        log_.WriteLine("An exception occurred in the [PdfDistiller.bSpoolJobs].")
                        Throw New ArgumentException(log_.ToString())
                    Case 0
                        log_.WriteLine("An exception occurred in the [PdfDistiller.FileToPDF()] method. Please check the argument.")
                        Throw New ArgumentException(log_.ToString())
                    Case 1
                        '' OK
                        log_.WriteLine("Successful in the creation of [" + outPDF + "] pdf file.")
                    Case Else
                        log_.WriteLine("It failed to execute [PdfDistiller.FileToPDF()] method.")
                        Throw New Exception(log_.ToString())
                End Select


                result(0) = outPDF
                result(1) = log_.ToString()
            Finally
            End Try
            Return result
        End Function
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
        End Sub

#Region "IDisposable Support "
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

#End Region
    End Class
End Namespace




