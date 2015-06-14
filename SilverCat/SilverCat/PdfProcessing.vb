Imports Acrobat
Imports ACRODISTXLib
Imports AFORMAUTLib
Imports System.IO
Imports System.Reflection


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

            Dim result As String() = New String() {Nothing, Nothing}
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
        ''' <paramref name="jsonSecurityParam">セキュリティ設定Javascriptに渡す引数。JSON形式のString。</paramref>
        ''' <remarks>
        ''' %Acrobatインストールディレクトリ%Acrobat\Javascriptsに、SetSecurity.jsを事前に配置のこと。
        ''' </remarks>
        ''' <returns>
        ''' String(0):出力PDFファイルのフルパス。
        ''' String(1):ログメッセージ文字列。
        ''' </returns>
        Public Function SignPdf(ByVal inPdfFilePath As String, _
                                ByVal outPdfFilePath As String, _
                                ByVal jsonSecurityParam As String) As String()

            Dim inPdDoc As AcroPDDoc = Nothing
            Dim inAvDoc As AcroAVDoc = Nothing
            Dim formApp As AFormApp = Nothing
            Dim fields As Fields = Nothing
            Dim rc As Boolean = False

            Dim result As String() = New String() {Nothing, Nothing}
            Try
                log_.Write(">>Input Pdf file is (")
                log_.Write(inPdfFilePath)
                log_.WriteLine(")")

                '' 電子署名する前の入力PDFファイルのまま残しておきたいので、コピーする。
                File.Copy(inPdfFilePath, outPdfFilePath, True)

                inPdDoc = New AcroPDDoc()
                rc = inPdDoc.Open(outPdfFilePath)
                If Not (rc) Then
                    Throw New IOException(">>It failed to open, for file(" & outPdfFilePath & ") ")
                End If

                '' Acrobat Proを立ち上げて、その中にPDFファイルを開く
                inAvDoc = CType(inPdDoc.OpenAVDoc("TEMP_PDF_DOCUMENT"), AcroAVDoc)

                '' 電子署名自体はJavascriptで行うので、ここではJavascriptを呼び出すイメージ。
                formApp = New AFormApp()

                fields = CType(formApp.Fields, Fields)

                Dim nVersion As String
                nVersion = fields.ExecuteThisJavascript("event.value = app.viewerVersion;")
                log_.WriteLine("The Acrobat viewer version is " & nVersion & ".")

                '' 電子署名javascriptを実行します。
                Dim jsCode As New StringWriter
                jsCode.Write("SignToPdf(this, ")
                jsCode.Write(jsonSecurityParam)
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

                '' javascriptの中で電子署名した時点でPDFファイルの保存をするので、ここでは保存せずに閉じます。
                '' というか、javascriptの外側では保存できないので、こうせざるを得ない。
                rc = inAvDoc.Close(1)

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

#Region "PDFファイルにセキュリティを設定します。"
        ''' <summary>
        ''' PDFファイルにセキュリティを設定します。
        ''' </summary>
        ''' <paramref name="inPdfFilePath">入力PDFファイルへのフルパス。</paramref>
        ''' <paramref name="outPdfFilePath">出力PDFファイルへのフルパス。</paramref>
        ''' <paramref name="jsonSecurityParam">セキュリティ設定Javascriptに渡す引数。JSON形式のString。</paramref>
        ''' <remarks>
        ''' %Acrobatインストールディレクトリ%Acrobat\Javascriptsに、SetSecurity.jsを事前に配置のこと。
        ''' </remarks>
        ''' <returns>
        ''' String(0):出力PDFファイルのフルパス。
        ''' String(1):ログメッセージ文字列。
        ''' </returns>
        Public Function SetSecurityPdf(ByVal inPdfFilePath As String, _
                                ByVal outPdfFilePath As String, _
                                ByVal jsonSecurityParam As String) As String()

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

                inPdDoc = New AcroPDDoc()
                rc = inPdDoc.Open(inPdfFilePath)
                If Not (rc) Then
                    Throw New IOException(">>It failed to open, for file(" & inPdfFilePath & ") ")
                End If

                '' Acrobat Proを立ち上げて、その中にPDFファイルを開く
                inAvDoc = CType(inPdDoc.OpenAVDoc("TEMP_PDF_DOCUMENT"), AcroAVDoc)

                '' セキュリティ設定はJavascriptで行うので、ここではJavascriptを呼び出すイメージ。
                formApp = New AFormApp()

                fields = CType(formApp.Fields, Fields)

                Dim nVersion As String
                nVersion = fields.ExecuteThisJavascript("event.value = app.viewerVersion;")
                log_.WriteLine("The Acrobat viewer version is " & nVersion & ".")


                '' セキュリティ設定javascriptを実行します。
                Dim jsCode As New StringWriter
                jsCode.Write("SetSecurityToPdf(this, ")
                jsCode.Write(jsonSecurityParam)
                jsCode.Write(",'")
                jsCode.Write(outPdfFilePath)
                jsCode.Write("');")
                Dim jsRc As String = fields.ExecuteThisJavascript(jsCode.ToString())

                '' 戻り値:ゼロが正常終了。それ以外は、異常終了。
                If (jsRc <> "0") Then
                    Throw New Exception("SetSecurityToPdf:" & jsRc)
                End If

                '' AVDoc.Close(bNoSave)
                '' bNoSave:
                '' 正の数:PDFドキュメントを保存しないで閉じます。 
                '' 0でPDFドキュメントが変更されていた:保存するかどうか確認するダイアログを表示
                '' 負の数:確認なしにPDFドキュメントを保存。

                '' javascriptの中でPDFファイルの保存をするので、ここでは保存せずに閉じます。
                '' というか、javascriptの外側では保存できないので、こうせざるを得ない。
                rc = inAvDoc.Close(1)

                If rc Then
                    log_.WriteLine(">>security pdf document to save (" & outPdfFilePath & ").")
                Else
                    Throw New IOException("It failed security pdf document to save (" + outPdfFilePath + ").")
                End If

                result(0) = outPdfFilePath
                result(1) = log_.ToString()

            Catch ex As Exception
                '' セキュリティ設定用に作ったPDFドキュメントを保存せずに閉じます。
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

#Region "PDFファイルを加工指示命令に従って加工します。"
        ''' <summary>
        ''' PDFファイルを加工指示命令に従って加工します。
        ''' </summary>
        ''' <paramref name="inPdfFilePath">入力PDFファイルへのフルパス。</paramref>
        ''' <paramref name="outPdfFilePath">出力PDFファイルへのフルパス。</paramref>
        ''' <paramref name="processingParameter">PDFファイル加工指示命令。JSON形式のString。</paramref>
        ''' <remarks>
        ''' </remarks>
        ''' <returns>
        ''' String(0):出力PDFファイルのフルパス。
        ''' String(1):ログメッセージ文字列。
        ''' </returns>
        Public Function ProcessPdf(ByVal inPdfFilePath As String, _
                                ByVal outPdfFilePath As String, _
                                ByVal processingParameter As Dictionary(Of String, Dictionary(Of String, Object))) As String()

            Dim inPdDoc As AcroPDDoc = Nothing
            Dim jsObj As JSObject = Nothing
            Dim rc As Boolean

            Dim result() As String = New String() {Nothing, Nothing}
            Try
                log_.Write(">>Input Pdf file is (")
                log_.Write(inPdfFilePath)
                log_.WriteLine(")")

                '' ウォーターマークPDFファイルを開きます。
                inPdDoc = New AcroPDDoc()
                rc = inPdDoc.Open(inPdfFilePath)
                If Not (rc) Then
                    Throw New IOException(">>It failed to open, for file(" & inPdfFilePath & ") ")
                End If

                '' JSONオブジェクトの取得。
                jsObj = New JSObject(inPdDoc.GetJSObject())

                '' 前面にウォーターマークテキストを付与します。
                Dim foreground As Dictionary(Of String, Object) = CType(processingParameter.Item("foreground"), Dictionary(Of String, Object))
                If Not (foreground Is Nothing) Then
                    jsObj.addWatermarkFromText(foreground)
                End If

                '' ウォーターマークPDFファイルによりウォーターマークを付与します。
                Dim watermark As Dictionary(Of String, Object) = CType(processingParameter.Item("waterMark"), Dictionary(Of String, Object))
                If Not (watermark Is Nothing) Then
                    jsObj.addWatermarkFromFile(watermark)
                End If

                '' XMPメタを付与します。
                Dim xmp As Dictionary(Of String, Object) = CType(processingParameter.Item("xmp"), Dictionary(Of String, Object))
                If Not (xmp Is Nothing) Then
                    Dim xmpString As Object
                    xmpString = System.IO.File.ReadAllText(xmp.Item("path"), System.Text.Encoding.UTF8)
                    jsObj.Metadata = xmpString
                End If

                '' 出来上がったPDFファイルをファイルとしてに保存します。
                rc = inPdDoc.Save(PDSaveFlags.PDSaveFull, outPdfFilePath)
                If Not (rc) Then
                    Throw New IOException(">>It failed to add watermark, for file(" & inPdfFilePath & ") ")
                End If

                result(0) = outPdfFilePath
                result(1) = log_.ToString()
            Finally
                If Not (jsObj Is Nothing) Then
                    jsObj = Nothing
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

#Region "JSONオブジェクト"
        ''' <summary>
        ''' Adobe Acrobat JSONオブジェクトのラッパークラス
        ''' </summary>
        Class JSObject
            Implements IDisposable
            ''' <summary>Acrobat JSONオブジェクトのインスタンス。</summary>
            Private acroJson_ As Object

            ''' <summary>Acrobat Colorインスタンス。</summary>
            Private acroColor_ As Object

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="acroJson">Acrobat JSONオブジェクト。</param>
            Public Sub New(ByRef acroJson As Object)
                Me.acroJson_ = acroJson
                Dim jsonType As Type = Me.acroJson_.GetType()
                Me.acroColor_ = jsonType.InvokeMember("color", BindingFlags.GetProperty Or BindingFlags.Public Or BindingFlags.Instance, Nothing, Me.acroJson_, Nothing)
            End Sub
            ''' <summary>XMPメタデータ。</summary>
            Public Property Metadata() As Object
                Set(value As Object)
                    Me.acroJson_.GetType().InvokeMember("metadata", BindingFlags.SetProperty Or BindingFlags.Public Or BindingFlags.Instance, Nothing, Me.acroJson_, New Object() {value})
                End Set
                Get
                    Return Me.acroJson_.GetType().InvokeMember("metadata", BindingFlags.GetProperty Or BindingFlags.Public Or BindingFlags.Instance, Nothing, Me.acroJson_, Nothing)
                End Get
            End Property

            '''　<summary>PDFファイルにPDFファイルをウォーターマークとして付け加えます。</summary>
            ''' <param name="waterMarkParam">
            ''' <code>
            ''' Dim addFileWatermarkParam As Object() = new Object() {
            '''     waterMarkFilePath, '' cDIPath:ウォーターマークPDFファイルパス。
            '''     0,       '' nSourcePage:ゼロは、ウォーターマークPDFファイルの最初のページをウォーターマークにする。
            '''     -1,      '' nStart:開始ページ。
            '''     -1,      '' nEnd:終了ページ。開始と終了ページを-1にするとすべてのページにウォーターマークを付与。
            '''     false,   '' bOnTop:ウォーターマークの前景と背景。true:前景、false:背景。
            '''     true,    '' bOnScreen:ウォーターマークをスクリーン表示する。true:表示する、false:表示しない。
            '''     true,    '' bOnPrint:ウォーターマークを印刷表示する。true:表示する、false:表示しない。
            '''     jsObj.AlignCenter,       '' nHorizAlign:ウォーターマークの配置:left:0, center:1, right:2, top:3, bottom:4
            '''     jsObj.AlignCenter,       '' nVartAlign:ウォーターマークの配置:left:0, center:1, right:2, top:3, bottom:4
            '''     0,       '' nHorizValue:左からの位置。XY 座標は（0,0）の左下隅
            '''     0,       '' nVartValue:下からの位置。XY 座標は（0,0）の左下隅
            '''     false,   '' bPercentage:上下左右からの位置をパーセントで指定する。true:パーセント指定する。false:パーセント指定しない。
            '''     -1.0,    '' nScale:ページに合わせた相対倍率のこと。1.0=100%でオリジナルのフォントサイズ。-1.0でページにフィット。
            '''     true,    '' bFixedPrint:ページサイズが異なる場合、ウォーターマーク位置とサイズを一定にする。true:一定にする。false:一定にしない。
            '''     0,       '' nRotation:ゼロ:0=回転ゼロ度。つまり回転しない。
            '''     0.1      '' nOpactiy:透明度:0.1=透明度10%。
            ''' };
            ''' </code>
            ''' </param>
            Public Sub addWatermarkFromFile(ByRef waterMarkParam As Dictionary(Of String, Object))
                Dim addFileWatermarkParam As Object() = New Object() {
                    waterMarkParam.Item("cDIPath"),
                    waterMarkParam.Item("nSourcePage"),
                    waterMarkParam.Item("nStart"),
                    waterMarkParam.Item("nEnd"),
                    waterMarkParam.Item("bOnTop"),
                    waterMarkParam.Item("bOnScreen"),
                    waterMarkParam.Item("bOnPrint"),
                    Me.ChangeAlignType(waterMarkParam.Item("nHorizAlign")),
                    Me.ChangeAlignType(waterMarkParam.Item("nVertAlign")),
                    waterMarkParam.Item("nHorizValue"),
                    waterMarkParam.Item("nVertValue"),
                    waterMarkParam.Item("bPercentage"),
                    CType(waterMarkParam.Item("nScale"), Double),
                    waterMarkParam.Item("bFixedPrint"),
                    waterMarkParam.Item("nRotation"),
                    CType(waterMarkParam.Item("nOpacity"), Double)
                }
                ' Add a watermark from a file.
                ' function prototype:
                '   addWatermarkFromFile(cDIPath, nSourcePage, nStart, nEnd, bOnTop, bOnScreen, bOnPrint, nHorizAlign, nVertAlign, nHorizValue, nVertValue, bPercentage, nScale, bFixedPrint, nRotation, nOpacity)
                Me.acroJson_.GetType().InvokeMember("addWatermarkFromFile", BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance, Nothing, Me.acroJson_, addFileWatermarkParam)
            End Sub
            ''' <summary>
            ''' 文字列をウォーターマークとしてPDFファイルに付与します。
            ''' </summary>
            ''' <param name="waterMarkParam">
            ''' <code>
            ''' Dim addFileWatermarkParam As Object() = new Object() {
            '''     "COPY",             '' cText:ウォーターマーク文字列
            '''     jsObj.AlignCenter,  '' nTextAlign
            '''     "MS-Gothic",        '' cFont:フォント名。%Acrobatインストールディレクトリ%Resource\CIDFontの下にあるフォントとか。
            '''                         '' PostScriptファイルで指定する形式のフォント名で記述する。
            '''     100,                '' nFontSize:フォントサイズ(単位ポイント)。1 pt = 1/72 in. (= 25.4/72 mm = 0.352 777 7... mm)。100=100pt。
            '''     jsObj.Blue,         '' aColor:文字色。"black","blue","cyan","dkGray","gray","green","ltGray","magenta","red","white","yellow"。
            '''     0,                  '' nStart:開始ページ。
            '''     0,                  '' nEnd:終了ページ。開始と終了ページを-1にするとすべてのページにウォーターマークを付与。
            '''     true,               '' bOnTop:ウォーターマークの前景と背景。true:前景、false:背景。
            '''     true,               '' bOnScreen:ウォーターマークをスクリーン表示する。true:表示する、false:表示しない。
            '''     true,               '' bOnPrint:ウォーターマークを印刷表示する。true:表示する、false:表示しない。
            '''     jsObj.AlignCenter,  '' nHorizAlign:ウォーターマークの配置:left:0, center:1, right:2, top:3, bottom:4
            '''     jsObj.AlignTop,     '' nVartAlign:ウォーターマークの配置:left:0, center:1, right:2, top:3, bottom:4
            '''     20,                 '' nHorizValue:左からの位置。XY 座標は（0,0）の左下隅
            '''     -45,                '' nVartValue:下からの位置。XY 座標は（0,0）の左下隅
            '''     false,              '' bPercentage:上下左右からの位置をパーセントで指定する。true:パーセント指定する。false:パーセント指定しない。
            '''     1.0,                '' nScale:ページに合わせた相対倍率のこと。1.0=100%でオリジナルのフォントサイズ。-1.0でページにフィット。
            '''     false,              '' bFixedPrint:ページサイズが異なる場合、ウォーターマーク位置とサイズを一定にする。true:一定にする。false:一定にしない。
            '''     0,                  '' nRotation:ゼロ:0=回転ゼロ度。つまり回転しない。
            '''     0.7                 '' nOpactiy:透明度:0.7=透明度70%
            ''' };
            ''' </code>
            ''' </param>
            Public Sub addWatermarkFromText(ByRef waterMarkParam As Dictionary(Of String, Object))
                Dim addTextWatermarkParam As Object() = New Object() {
                    waterMarkParam.Item("cText"),
                    Me.ChangeAlignType(waterMarkParam.Item("nTextAlign")),
                    waterMarkParam.Item("cFont"),
                    waterMarkParam.Item("nFontSize"),
                    Me.GetColor(waterMarkParam.Item("aColor")),
                    waterMarkParam.Item("nStart"),
                    waterMarkParam.Item("nEnd"),
                    waterMarkParam.Item("bOnTop"),
                    waterMarkParam.Item("bOnScreen"),
                    waterMarkParam.Item("bOnPrint"),
                    Me.ChangeAlignType(waterMarkParam.Item("nHorizAlign")),
                    Me.ChangeAlignType(waterMarkParam.Item("nVertAlign")),
                    waterMarkParam.Item("nHorizValue"),
                    waterMarkParam.Item("nVertValue"),
                    waterMarkParam.Item("bPercentage"),
                    CType(waterMarkParam.Item("nScale"), Double),
                    waterMarkParam.Item("bFixedPrint"),
                    waterMarkParam.Item("nRotation"),
                    CType(waterMarkParam.Item("nOpacity"), Double)
                }
                Me.acroJson_.GetType().InvokeMember("addWatermarkFromText", BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance, Nothing, Me.acroJson_, addTextWatermarkParam)
            End Sub

            ''' <summary>色の取得</summary>
            ''' <param name="colorName">取得したい色の名称</param>
            ''' <remarks>
            ''' 使える色の名前は次の名前。
            ''' "black","blue","cyan","dkGray","gray","green","ltGray","magenta","red","white","yellow"
            ''' </remarks>
            Public Function GetColor(ByRef colorName As String) As Object
                Return Me.acroJson_.GetType().InvokeMember(colorName, BindingFlags.GetProperty Or BindingFlags.Public Or BindingFlags.Instance, Nothing, Me.acroColor_, Nothing)
            End Function
            ''' <summary>整列位置の取得</summary>
            ''' <param name="alignName">取得したい位置の名称</param>
            ''' <remarks>
            ''' 使える位置の名前は次の名前。
            ''' "left","center","right","top","bottom"
            ''' </remarks>
            Public Function ChangeAlignType(ByVal alignName As String) As Integer
                Dim align As Integer
                Select Case alignName
                    Case "left"
                        align = Me.acroJson_.app.constants.align.left
                    Case "center"
                        align = Me.acroJson_.app.constants.align.center
                    Case "right"
                        align = Me.acroJson_.app.constants.align.right
                    Case "top"
                        align = Me.acroJson_.app.constants.align.top
                    Case "bottom"
                        align = Me.acroJson_.app.constants.align.bottom
                    Case Else
                        Throw New ArgumentException(">>ChangeAlign Function argment (" & alignName & ") does not work.")
                End Select
                Return align
            End Function
            ''' <summary>リソースの解放を行います。</summary>
            Protected disposed As Boolean = False
            Protected Overridable Sub Dispose(ByVal disposing As Boolean)
                If Not Me.disposed Then
                    If Not (Me.acroColor_ Is Nothing) Then
                        Marshal.ReleaseComObject(Me.acroColor_)
                        Me.acroColor_ = Nothing
                    End If
                    If Not (Me.acroJson_ Is Nothing) Then
                        Marshal.ReleaseComObject(Me.acroJson_)
                        Me.acroJson_ = Nothing
                    End If
                End If
            End Sub
            ''' <summary>リソースの解放を行います。</summary>
            Public Sub Dispose() Implements IDisposable.Dispose
                Dispose(True)
                GC.SuppressFinalize(Me)
            End Sub
            Protected Overrides Sub Finalize()
                Dispose(False)
                MyBase.Finalize()
            End Sub
        End Class
#End Region

#Region "Dispose"
        Protected disposed As Boolean = False
        ''' <summary>
        ''' リソース解放
        ''' </summary>
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            Console.Error.WriteLine(">PdfProcessing:Resources of the release.")
            If Not Me.disposed Then
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
            End If
            Me.disposed = True
        End Sub
#End Region
#Region "IDisposable Support "
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overrides Sub Finalize()
            Dispose(False)
            MyBase.Finalize()
        End Sub
#End Region


    End Class


End Namespace
