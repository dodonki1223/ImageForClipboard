Option Explicit On

Imports ImageForClipboard.ImageForClipboardDefinition

''' <summary>アプリケーションクラス</summary>
''' <remarks>
'''   アプリケーションクラスの機能
'''     アプリケーションの有効期間を追跡し、これと対話する。
'''     コマンド ライン パラメーターを取得し、処理する。
'''     未処理の例外を検出し、これに応答する。
'''     アプリケーション スコープのプロパティとリソースを共有する。
'''     スタンドアロン アプリケーションのウィンドウを管理する。
'''     ナビゲーションを追跡し管理する。
''' </remarks>
Class Application

    ' Startup、Exit、DispatcherUnhandledException などのアプリケーション レベルのイベントは、
    ' このファイルで処理できます。

#Region "列挙体"

    ''' <summary>エラーメッセージタイプ</summary>
    ''' <remarks></remarks>
    Public Enum ErrorMessageType

        ''' <summary>メッセージボックス</summary>
        MessageBox

        ''' <summary>イベントログ</summary>
        EventLog

    End Enum

#End Region

#Region "コンストラクタ"

    ''' <summary>コンストラクタ</summary>
    ''' <remarks></remarks>
    Public Sub New()

        ' 例外が処理されなかったら発生するイベントを設定
        AddHandler Me.DispatcherUnhandledException, AddressOf Application_DispatcherUnhandledException

    End Sub

#End Region

#Region "イベント"

    ''' <summary>
    '''   アプリケーションのスタートアップイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Application_Startup(ByVal sender As System.Object, ByVal e As System.Windows.StartupEventArgs)

        'コマンドライン引数情報を取得
        Dim mCommandLine As New CommandLine(e.Args)

        'メインウィンドウのプロパティをセット
        Dim mMainWindow As New MainWindow()
        _SetMainWindowProperty(mMainWindow, mCommandLine)

        'ヘルプコマンドが存在したら、処理を終了
        If mCommandLine.IsExistsHelpCommand Then Me.Shutdown()

        'コマンドライン引数が不正だった時
        If Not mCommandLine.CommandLineException Is Nothing Then

            'コマンドラインクラスで取得した例外メッセージを表示
            MessageBox.Show(mCommandLine.CommandLineException.Message _
                          , cNameSpaceName _
                          , MessageBoxButton.OK _
                          , MessageBoxImage.Error)

            '処理を終了
            Me.Shutdown()

        End If

        'メインウィンドウを表示
        mMainWindow.Show()

    End Sub

    ''' <summary>未処理例外をキャッチするイベントハンドラ</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    '''   「try ～ Catch」で補足されなかった例外をキャッチします。
    '''   このイベントは、キャッチされない例外の通知を提供します。 
    '''   これにより、アプリケーションをシステムの既定のハンドラーがユーザーに例外を報告し、
    '''   アプリケーションを終了する前に、例外に関する情報を記録できます。
    '''   その他の操作を行うことがあります、アプリケーションの状態に関する十分な情報が利用可能な場合は、
    '''   -など、その後の復旧のプログラム データを保存します。 
    '''   注意が必要、例外が処理されない場合に、プログラムのデータが破壊されることがあるためです。
    ''' </remarks>
    Private Sub Application_DispatcherUnhandledException(ByVal sender As Object, ByVal e As System.Windows.Threading.DispatcherUnhandledExceptionEventArgs)

        'Exceptionクラスに変換する（できなかった場合はNothingが返る）
        Dim mException As Exception = TryCast(e.Exception, Exception)

        'ExceptionがNothingの時
        If mException Is Nothing Then

            MessageBox.Show("System.Exceptionとして扱えない例外です", "", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub

        End If

        '------------------------
        ' メッセージボックス表示
        '------------------------
        'メッセージボックスに表示するメッセージを取得
        Dim mMessgeBoxMessage As String = GetErrorMessage(mException, ErrorMessageType.MessageBox)

        'メッセージボックスを表示
        MessageBox.Show(mMessgeBoxMessage, "UnhandledException", MessageBoxButton.OK, MessageBoxImage.Error)

        '------------------------
        ' イベントログに書き込み
        '------------------------
        'イベントログに表示するメッセージを取得
        Dim mEventLogMessage As String = GetErrorMessage(mException, ErrorMessageType.EventLog)

        'イベントログに書き込む
        Call WriteEventLog(mException, mEventLogMessage, EventLogEntryType.Error)

        'このプロセスを終了し、オペレーティング システムに終了コードを返す
        '※処理が正常に完了したことを示す場合は 0 (ゼロ) を使用します
        Environment.Exit(0)

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>
    '''   コマンドライン引数からメインウィンドウのプロパティをセット
    ''' </summary>
    ''' <param name="pWindow">メインウィンドウ</param>
    ''' <param name="pCommandLine">コマンドラインクラス</param>
    ''' <remarks></remarks>
    Private Sub _SetMainWindowProperty(ByVal pWindow As MainWindow, ByVal pCommandLine As CommandLine)

        'ウィンドウを自動で閉じる時間プロパティが存在する時はセット
        If Not pCommandLine.AutoCloseTime = Nothing Then pWindow.AutoCloseTime = pCommandLine.AutoCloseTime

        '画像の保存先パスプロパティが存在する時
        If Not pCommandLine.SaveImagePath Is Nothing Then

            '画像の保存先パスの末尾が「\」バックスラッシュまたは「/」スラッシュでない時は「\」バックスラッシュを末尾に追加
            Dim mSavePath As String = pCommandLine.SaveImagePath
            If mSavePath.Substring(pCommandLine.SaveImagePath.Length - 1, 1) <> "\" _
            AndAlso mSavePath.Substring(pCommandLine.SaveImagePath.Length - 1, 1) <> "/" Then mSavePath = mSavePath & "\"

            '画像の保存先パスプロパティをセット
            pWindow.SaveImagePath = mSavePath

            '画像の保存ファイル名プロパティをセット ※ファイル名は「西暦 + 月 + 日 + 時間 + 分 + 秒」とする
            pWindow.SaveImageFileName = DateTime.Now.ToString("yyyyMMddHHmmss")

        Else

            '画像の保存ファイル名プロパティをセット（デフォルト保存ファイル名）
            pWindow.SaveImageFileName = cDefaultSaveImageFileName

        End If

        '出力画像拡張子プロパティをセット
        pWindow.OutputImageExtension = pCommandLine.OutputImageExtension

        'クリップボードにコピーする画像パスプロパティが存在する時はセット
        If Not pCommandLine.CopyImageToClipboardPath = Nothing Then pWindow.CopyImageToClipboardPath = pCommandLine.CopyImageToClipboardPath

        '画像表示サイズプロパティをセット
        pWindow.DisplayImageSize = pCommandLine.DisplayImageSize

        'ウィンドウを非表示するかどうかプロパティがTrueのときはセット
        If pCommandLine.DoNotShowWindow = True Then pWindow.DoNotShowWindow = pCommandLine.DoNotShowWindow

    End Sub

    ''' <summary>エラーメッセージを取得する</summary>
    ''' <param name="pEx">Exceptionクラス</param>
    ''' <param name="pMessageType">エラーメッセージタイプ</param>
    ''' <returns>エラーメッセージ</returns>
    ''' <remarks>エラーメッセージタイプに応じたエラーメッセージを返す</remarks>
    Public Shared Function GetErrorMessage(ByVal pEx As Exception, ByVal pMessageType As ErrorMessageType) As String

        Dim mErrorMessage As New System.Text.StringBuilder

        'エラーメッセージタイプにより処理を分岐
        Select Case pMessageType

            Case ErrorMessageType.MessageBox

                With mErrorMessage

                    .AppendLine("エラーが発生しました。開発元にお知らせ下さい。")
                    .AppendLine()
                    .AppendLine("【エラー内容】")
                    .AppendLine(pEx.Message)

                End With

            Case ErrorMessageType.EventLog

                With mErrorMessage

                    .AppendLine("【エラー内容】")
                    .AppendLine(pEx.Message)
                    .AppendLine()
                    .AppendLine("【エラー発生メソッド】")
                    .AppendLine("「" & pEx.TargetSite.Name & "」メソッドでエラーが発生")
                    .AppendLine()
                    .AppendLine("【スタックトレース】")
                    .AppendLine(pEx.StackTrace)

                End With

        End Select

        Return mErrorMessage.ToString

    End Function

    ''' <summary>イベントログに書き込む</summary>
    ''' <param name="pEx">Exceptionクラス</param>
    ''' <param name="pEntryType">イベントログエントリのイベントの種類</param>
    ''' <remarks>
    '''   ユーザーアカウント制御 （UAC）が有効の場合、「管理者として実行」しないと、
    '''   例外（「System.Security.SecurityException: 要求されたレジストリ アクセスは許可されていません。」）が発生しますので、可能ならばVSを管理者として実行してください。
    '''   「If Not EventLog.SourceExists(mSourceName) Then EventLog.CreateEventSource(mSourceName, "")」この行で上記のようなエラーが発生した
    '''   Visual Studioは管理者で実行すること
    ''' </remarks>
    Public Shared Sub WriteEventLog(ByVal pEx As Exception, ByVal pErrorMessage As String, ByVal pEntryType As EventLogEntryType)

        'ソース名を取得
        Dim mSourceName As String = pEx.Source

        'ソースが存在していない時は作成する
        If Not EventLog.SourceExists(mSourceName) Then EventLog.CreateEventSource(mSourceName, "")

        'イベントログにエントリを書き込む
        EventLog.WriteEntry(mSourceName, pErrorMessage, pEntryType)

    End Sub

#End Region

End Class
