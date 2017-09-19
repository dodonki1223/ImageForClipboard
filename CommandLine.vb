Option Explicit On

Imports ImageForClipboard.ImageForClipboardDefinition

''' <summary>
'''   コマンドライン引数の機能を提供する
''' </summary>
''' <remarks>
'''   コマンドライン引数からコマンドラインキーごと分割した値を保持するクラスです
''' </remarks>
Public Class CommandLine

#Region "定数"

    ''' <summary>
    '''   メッセージ定数
    ''' </summary>
    ''' <remarks>CommandLineクラスで使用するメッセージを提供します</remarks>
    Private Class _cMessage

        ''' <summary>「コマンドライン引数が不正です」メッセージ</summary>
        Public Const InvalidCommandLineArgs As String = "コマンドライン引数が不正です"

        ''' <summary>「ウィンドウを自動で閉じる時間が不正です」メッセージ</summary>
        Public Const InvalidAutoCloseTime As String = "ウィンドウを自動で閉じる時間が不正です"

        ''' <summary>「画像の保存先パスが不正です」メッセージ</summary>
        Public Const InvalidSaveImagePath As String = "画像の保存先パスが不正です"

        ''' <summary>「クリップボードにコピーする画像パスが不正です」メッセージ</summary>
        Public Const InvalidCopyImageToClipboardPath As String = "クリップボードにコピーする画像パスが不正です"

        ''' <summary>「出力画像拡張子が不正です」メッセージ</summary>
        Public Const InvalidOutputaImageExtension As String = "出力画像拡張子が不正です"

        ''' <summary>「表示画像サイズの指定が不正です」メッセージ</summary>
        Public Const InvalidDisplayImageSize As String = "表示画像サイズの指定が不正です"

    End Class

    ''' <summary>
    '''   コマンドライン引数コマンドキー定数
    ''' </summary>
    ''' <remarks></remarks>
    Private Class _cCommadLineKey

        ''' <summary>ヘルプ</summary>
        ''' <remarks>コマンドライン引数のヘルプを表示</remarks>
        Public Const Help As String = "/?"

        ''' <summary>ウィンドウを自動で閉じる</summary>
        ''' <remarks>時間（秒）を指定してウィンドウを閉じる</remarks>
        Public Const AutoClose As String = "/AutoClose"

        ''' <summary>画像を自動保存する</summary>
        ''' <remarks>対象パスへクリップボードの画像を保存する</remarks>
        Public Const AutoSave As String = "/AutoSave"

        ''' <summary>画像をクリップボードへコピー</summary>
        ''' <remarks>対象画像パスをクリップボードへコピーする</remarks>
        Public Const Copy As String = "/Copy"

        ''' <summary>ウィンドウを表示しない</summary>
        ''' <remarks></remarks>
        Public Const DoNotShow As String = "/DoNotShow"

        ''' <summary>出力画像拡張子</summary>
        ''' <remarks></remarks>
        Public Const Extension As String = "/Extension"

        ''' <summary>表示画像サイズ</summary>
        ''' <remarks>１～１００のみを指定すること</remarks>
        Public Const DisplayImageSize As String = "/ImageSize"

    End Class

#End Region

#Region "列挙体"

    ''' <summary>コマンドライン引数タイプ</summary>
    ''' <remarks></remarks>
    Private Enum _CommandLineType

        ''' <summary>キーのみ</summary>
        OnlyKey

        ''' <summary>キーと値</summary>
        KeyAndValue

    End Enum

#End Region

#Region "コンストラクタ"

    ''' <summary>
    '''   コンストラクタ
    ''' </summary>
    ''' <remarks>
    '''   引数なしコンストラクタは外部には公開しない
    ''' </remarks>
    Private Sub New()

    End Sub

    ''' <summary>
    '''   引数付きコンストラクタ
    ''' </summary>
    ''' <param name="pCommandLineArgs">コマンドライン引数配列</param>
    ''' <remarks>
    '''   コマンドライン引数を分割して変数にセットする
    ''' </remarks>
    Public Sub New(ByVal pCommandLineArgs As String())

        '------------------------------------------
        ' ヘルプを取得
        '------------------------------------------
        'ヘルプコマンドが存在した時
        If _GetCommandLineValue(pCommandLineArgs, _cCommadLineKey.Help, "", _CommandLineType.OnlyKey) Then

            'コマンドリストを表示し処理を終了
            ShowCommnadList()
            _IsExistsHelpCommand = True
            Exit Sub

        End If

        '------------------------------------------
        ' ウィンドウを自動で閉じる時間を取得
        '------------------------------------------
        'ウィンドウが自動で閉じる時間の取得用変数
        Dim mAutoCloseTimeString As String = String.Empty

        'ウィンドウが自動で閉じる時間の取得が成功した時
        If _GetCommandLineValue(pCommandLineArgs, _cCommadLineKey.AutoClose, mAutoCloseTimeString, _CommandLineType.KeyAndValue) Then

            'ウィンドウが自動で閉じる時間が数値に変換できる時
            Dim mAutoCloseTime As Integer
            If Integer.TryParse(mAutoCloseTimeString, mAutoCloseTime) Then

                _AutoCloseTime = New TimeSpan(0, 0, 0, mAutoCloseTime)

            Else

                _CommandLineException = New ArgumentException(_cMessage.InvalidAutoCloseTime)

            End If

        End If

        '------------------------------------------
        ' 画像の保存先パスを取得
        '------------------------------------------
        Dim mSaveImagePath As String = String.Empty

        '画像の保存先パスの取得が成功した時
        If _GetCommandLineValue(pCommandLineArgs, _cCommadLineKey.AutoSave, mSaveImagePath, _CommandLineType.KeyAndValue) Then

            '画像の保存先パスが存在する時
            If System.IO.Directory.Exists(mSaveImagePath) Then

                _SaveImagePath = mSaveImagePath

            Else

                _CommandLineException = New ArgumentException(_cMessage.InvalidSaveImagePath)

            End If

        End If

        '------------------------------------------
        ' クリップボードにコピーする画像パスを取得
        '------------------------------------------
        Dim mCopyImageToClipboardPath As String = String.Empty

        'クリップボードにコピーする画像パスの取得が成功した時
        If _GetCommandLineValue(pCommandLineArgs, _cCommadLineKey.Copy, mCopyImageToClipboardPath, _CommandLineType.KeyAndValue) Then

            'クリップボードにコピーする画像パスが存在する時
            If System.IO.File.Exists(mCopyImageToClipboardPath) Then

                _CopyImageToClipboardPath = mCopyImageToClipboardPath

            Else

                _CommandLineException = New ArgumentException(_cMessage.InvalidCopyImageToClipboardPath)

            End If

        End If

        '------------------------------------------
        ' 出力画像拡張子を取得
        '------------------------------------------
        '出力画像拡張子の取得用変数
        Dim mOutputExtensionString As String = String.Empty

        '出力画像拡張子の取得が成功した時
        If _GetCommandLineValue(pCommandLineArgs, _cCommadLineKey.Extension, mOutputExtensionString, _CommandLineType.KeyAndValue) Then

            '取得した出力画像拡張子の値が「画像の出力形式」列挙体に変換できる時
            Dim mOutputExtension As OutputExtension
            If [Enum].TryParse(Of OutputExtension)(mOutputExtensionString.ToLower, mOutputExtension) Then

                _OutputImageExtension = mOutputExtension

            Else

                _CommandLineException = New ArgumentException(_cMessage.InvalidOutputaImageExtension)

            End If

        End If

        '------------------------------------------
        ' 表示画像サイズを取得
        '------------------------------------------
        Dim mDisplayImageSizeString As String = String.Empty

        '表示画像サイズの取得が成功した時
        If _GetCommandLineValue(pCommandLineArgs, _cCommadLineKey.DisplayImageSize, mDisplayImageSizeString, _CommandLineType.KeyAndValue) Then

            Dim mDisplayImageSize As Double

            '     取得した表示画像サイズがDoubleに変換出来
            'かつ その範囲が１から１００までの時
            If Double.TryParse(mDisplayImageSizeString, mDisplayImageSize) _
            AndAlso 1 <= mDisplayImageSize _
            AndAlso mDisplayImageSize <= 100 Then

                _DisplayImageSize = mDisplayImageSize

            Else

                _CommandLineException = New ArgumentException(_cMessage.InvalidDisplayImageSize)

            End If

        End If

        '------------------------------------------
        ' ウィンドウを非表示するかどうかを取得
        '------------------------------------------
        _DoNotShowWindow = _GetCommandLineValue(pCommandLineArgs, _cCommadLineKey.DoNotShow, "", _CommandLineType.OnlyKey)

    End Sub

#End Region

#Region "プロパティ"

    ''' <summary>
    '''   コマンドライン引数例外プロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property CommandLineException() As ArgumentException

        Get

            Return _CommandLineException

        End Get

    End Property

    ''' <summary>
    '''   ウィンドウを自動で閉じる時間プロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property AutoCloseTime As TimeSpan

        Get

            Return _AutoCloseTime

        End Get

    End Property

    ''' <summary>
    '''   画像の保存先パスプロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property SaveImagePath As String

        Get

            Return _SaveImagePath

        End Get

    End Property

    ''' <summary>
    '''   クリップボードにコピーする画像パスプロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property CopyImageToClipboardPath As String

        Get

            Return _CopyImageToClipboardPath

        End Get

    End Property

    ''' <summary>
    '''   ウィンドウを非表示するかどうかプロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property DoNotShowWindow As Boolean

        Get

            Return _DoNotShowWindow

        End Get

    End Property

    ''' <summary>
    '''   出力画像拡張子プロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property OutputImageExtension As OutputExtension

        Get

            Return _OutputImageExtension

        End Get

    End Property

    ''' <summary>
    '''   表示画像サイズプロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property DisplayImageSize As Double

        Get

            Return _DisplayImageSize

        End Get

    End Property

    ''' <summary>
    '''   ヘルプコマンドが存在したかどうか
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property IsExistsHelpCommand As Boolean

        Get

            Return _IsExistsHelpCommand

        End Get

    End Property

#End Region

#Region "変数"

    ''' <summary>
    '''   コマンドライン引数の例外
    ''' </summary>
    ''' <remarks></remarks>
    Private _CommandLineException As ArgumentException

    ''' <summary>
    '''   ウィンドウを自動で閉じる時間を保持する変数
    ''' </summary>
    ''' <remarks></remarks>
    Private _AutoCloseTime As TimeSpan

    ''' <summary>
    '''   画像の保存先パスを保持する変数
    ''' </summary>
    ''' <remarks></remarks>
    Private _SaveImagePath As String

    ''' <summary>
    '''   クリップボードにコピーする画像パスを保持する変数
    ''' </summary>
    ''' <remarks></remarks>
    Private _CopyImageToClipboardPath As String

    ''' <summary>
    '''   ウィンドウを非表示するかどうかを保持する変数
    ''' </summary>
    ''' <remarks></remarks>
    Private _DoNotShowWindow As Boolean

    ''' <summary>
    '''   出力画像拡張子を保持する変数
    ''' </summary>
    ''' <remarks>デフォルトは「bmp」</remarks>
    Private _OutputImageExtension As OutputExtension = OutputExtension.bmp

    ''' <summary>
    '''   表示画像サイズを保持する変数
    ''' </summary>
    ''' <remarks>デフォルトは「50」つまり５０％表示</remarks>
    Private _DisplayImageSize As Double = 50

    ''' <summary>
    '''   ヘルプコマンドが存在するか
    ''' </summary>
    ''' <remarks>デフォルトは「bmp」</remarks>
    Private _IsExistsHelpCommand As Boolean = False

#End Region

#Region "メソッド"

    ''' <summary>
    '''   コマンドライン引数の値を取得する
    ''' </summary>
    ''' <param name="pCommandArgs">コマンドライン引数（配列）</param>
    ''' <param name="pCommandKey">対象コマンドキー</param>
    ''' <param name="pCommandValue">コマンドキーに対応する値</param>
    ''' <param name="pCommandType">コマンドライン引数タイプ</param>
    ''' <returns>True：対象コマンドキーが存在する、False：対象コマンドキーが存在しない（値が不正）</returns>
    ''' <remarks>
    '''   コマンドライン引数タイプにより処理を分岐
    '''     Onlykey    ：コマンドのみの時は存在チェックのみを行う
    '''     KeyAndValue：コマンドと値の時は存在チェックと値を取得
    ''' </remarks>
    Private Function _GetCommandLineValue(ByVal pCommandArgs As String(), ByVal pCommandKey As String, ByRef pCommandValue As String _
                                        , ByVal pCommandType As _CommandLineType) As Boolean

        'コマンドラインキー文字列が存在する位置を取得
        Dim mIndex As Integer = Array.IndexOf(pCommandArgs, pCommandKey)

        'コマンドラインキー文字列が存在しなかった時はFalseを返す
        If mIndex = -1 Then Return False

        Select Case pCommandType

            Case _CommandLineType.OnlyKey

                Return True

            Case _CommandLineType.KeyAndValue

                Try

                    'コマンドラインキーの次のコマンドライン引数の値を取得
                    pCommandValue = pCommandArgs(mIndex + 1)
                    Return True

                Catch ex As Exception

                    _CommandLineException = New ArgumentException(_cMessage.InvalidCommandLineArgs)
                    Return False

                End Try

            Case Else

                Return False

        End Select

    End Function

    ''' <summary>
    '''   コマンドリストを表示
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub ShowCommnadList()

        'コマンドリストの文字列を作成する
        Dim mCommnadListString As New System.Text.StringBuilder

        With mCommnadListString

            .AppendLine("ヘルプ　　　　　　　　　　　：/?                              ")
            .AppendLine("　・コマンドの一覧を表示します。その他のコマンドが設定されて  ")
            .AppendLine("　　いた場合はすべて無視されヘルプコマンドだけ実行され、処理  ")
            .AppendLine("　　を終了します                                              ")
            .AppendLine("ウィンドウを自動で閉じる　　：/AutoClose                      ")
            .AppendLine("　・ウィンドウを表示後、自動で閉じる時間（秒）を指定できます  ")
            .AppendLine("　 ※例：/AutoClose 5　５秒後にウィンドウが自動で閉じる       ")
            .AppendLine("画像をクリップボードへコピー：/Copy                           ")
            .AppendLine("　・画像パスを指定することでその画像をクリップボードへコピー  ")
            .AppendLine("　　します。画像が存在しない場合はエラーが表示されます。      ")
            .AppendLine("　※例：/Copy C:\test.png                                     ")
            .AppendLine("画像を自動保存する　　　　　：/AutoSave                       ")
            .AppendLine("　・指定したディレクトリにクリップボード内の画像を保存します。")
            .AppendLine("　　保存されるファイル名は「西暦月日時間分秒＋拡張子」となり  ")
            .AppendLine("　　ます                                                      ")
            .AppendLine("　　保存ファイル名例：20170914121011.bmp                      ")
            .AppendLine("　※例：/AutoSave C:\test                                     ")
            .AppendLine("ウィンドウを表示しない　　　：/DoNotShow                      ")
            .AppendLine("　・起動した時、ウィンドウを表示させません。画像を自動保存す  ")
            .AppendLine("　　る時に併用して使用して下さい                              ")
            .AppendLine("　※例：/DoNotShow                                            ")
            .AppendLine("出力画像拡張子　　　　　　　：/Extension                      ")
            .AppendLine("　・画像を自動保存する時の拡張子を指定出来ます                ")
            .AppendLine("　　対応している拡張子：bmp,png,gif,jpg                       ")
            .AppendLine("　　何も指定しない場合は「bmp」となります                     ")
            .AppendLine("　※例：/Extension jpg                                        ")
            .AppendLine("表示画像サイズ　　　　　　　：/ImageSize                      ")
            .AppendLine("　・表示する画像のサイズを指定することができます（１～１００）")
            .AppendLine("　※例：/ImageSize 40　４０％の大きさで画像を表示             ")

        End With

        MessageBox.Show(mCommnadListString.ToString, "コマンドリスト", MessageBoxButton.OK, MessageBoxImage.Information)

        mCommnadListString = Nothing

    End Sub

#End Region

End Class
