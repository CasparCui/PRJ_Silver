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
        Private Shared WithEvents pdfDistiller_ As PdfDistiller = New PdfDistiller


        ''' <summary>
        ''' Acrobat Distiller の ログ出力。
        ''' </summary>
        Private Shared distillerLog_ As StringWriter = New StringWriter

        ''' <summary>
        ''' コンストラクタ(不可視)。
        ''' </summary>
        Private Sub New()

        End Sub

        ''' <summary>
        ''' シングルトンを返します。
        ''' </summary>
        Public Shared Function GetInstance() As PdfProcessing
            Return me_
        End Function

        ''' <summary>
        ''' Acrobat Distillerの処理状態を管理するフラグ。
        ''' </summary>
        ''' <remarks>
        ''' Acrobat Distillerのイベント通知で、処理中ステータスの管理に利用。
        ''' TRUE:処理中。FALSE:アイドル。
        ''' </remarks>
        Private bWorking_ As Boolean

        ''' <summary>
        ''' PostScriptファイルからPdfを作成します。
        ''' </summary>
        Public Function CreatePdf(ByRef inPS As String, _
                             ByRef outPDF As String, _
                             ByRef inOptionFilePath As String) As String()

            Dim result() As String
            Try
                result = New String() {Nothing, Nothing}

                distillerLog_.Write(">>Input PostScript file is (")
                distillerLog_.Write(inPS)
                distillerLog_.WriteLine(").")

                AddHandler pdfDistiller_.OnJobDone, _
                    Sub(input As String, output As String)
                        distillerLog_.WriteLine(">>created the PDF(" & output & ").")
                    End Sub
                AddHandler pdfDistiller_.OnJobFail, _
                    Sub(input As String, output As String)
                        distillerLog_.WriteLine(">>failed to create the PDF(" & output & ").")
                    End Sub
                AddHandler pdfDistiller_.OnPercentDone, _
                    Sub(percentDone As Integer)
                        If percentDone > 0 Then
                            Me.bWorking_ = True
                        End If
                        If Me.bWorking_ Then
                            If percentDone = 0 Then
                                distillerLog_.WriteLine("Idle:")
                                Me.bWorking_ = False
                            Else
                                distillerLog_.WriteLine(percentDone & "% processing...")
                            End If
                        End If
                    End Sub
                AddHandler pdfDistiller_.OnPageNumber, _
                    Sub(pageNum As Integer)
                        distillerLog_.WriteLine("be creating (" & pageNum & ") page...")
                    End Sub

                pdfDistiller_.bShowWindow = False
                pdfDistiller_.bSpoolJobs = False
                Me.bWorking_ = False

                Dim rc As Integer

                rc = pdfDistiller_.FileToPDF(inPS, outPDF, inOptionFilePath)

                Select rc
                    Case -1
                        distillerLog_.WriteLine("PdfDistiller.bSpoolJobsフラグの設定に問題があります。")
                        Throw New ArgumentException(distillerLog_.ToString())
                    Case 0
                        distillerLog_.WriteLine("PdfDistiller.FileToPDF()メソッドの引数に問題があります。")
                        Throw New ArgumentException(distillerLog_.ToString())
                    Case 1
                        '' OK
                        distillerLog_.WriteLine("出力PDFファイル(" + outPDF + ")としてPDFファイルの作成に成功しました。")
                    Case Else
                        distillerLog_.WriteLine("PdfDistiller.FileToPDF()メソッドの実行に失敗しました。")
                        Throw New Exception(distillerLog_.ToString())
                End Select


                result(0) = outPDF
                result(1) = distillerLog_.ToString()
            Finally
            End Try
            Return result
        End Function
#Region "Dispose"
        ''' <summary>
        ''' リソース解放
        ''' </summary>
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            Console.Error.WriteLine(">PdfProcessing:終了処理。")
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




