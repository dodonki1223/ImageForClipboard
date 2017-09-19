Imports System.Windows.Interop
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports ImageForClipboard.ImageForClipboardDefinition

''' <summary>
'''   クリップボード内の画像を表示する機能を提供するウィンドウです
''' </summary>
''' <remarks></remarks>
Class MainWindow

#Region "定数"

    ''' <summary>
    '''   メッセージ定数
    ''' </summary>
    ''' <remarks>CommandLineクラスで使用するメッセージを提供します</remarks>
    Private Class _cMessage

        ''' <summary>「画像の保存に失敗しました」メッセージ</summary>
        Public Const SaveImageFailure As String = "画像の保存に失敗しました"

        ''' <summary>「クリップボードへ画像のコピーが失敗しました」メッセージ</summary>
        Public Const CopyImageToClipboardFailure As String = "クリップボードへ画像のコピーが失敗しました"

        ''' <summary>「ファイルの種類に存在しない拡張子を指定して画像の保存はできません」メッセージ</summary>
        Public Shared NotExistsExtensionForSaveAsDialog As String = "ファイルの種類に存在しない拡張子を指定して" & System.Environment.NewLine & _
                                                                    "画像の保存はできません"

    End Class

    ''' <summary>
    '''  ウィンドウの表示位置（マウスからの距離）
    ''' </summary>
    Private Const _cDisplayWindowPlaceMargin As Double = 10

#End Region

#Region "列挙体"

    ''' <summary>
    '''   コンテキストメニューアイテム
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ContextMenuItem

        クリップボードの画像を保存

        閉じる

    End Enum

#End Region

#Region "プロパティ"

    ''' <summary>
    '''   クリップボード内の画像を取得プロパティ
    ''' </summary>
    ''' <remarks>
    '''   クリップボード内のデータがBimap形式のもので無い時はNothingを返す
    '''   Bitmap形式の時はBitmap画像を返す
    ''' </remarks>
    Public ReadOnly Property ClipboardImage As InteropBitmap

        Get

            'クリップボード内のデータを取得
            Dim mClipboardData As IDataObject = Clipboard.GetDataObject()

            'クリップボード内のデータがBitmapに変換出来る時
            If mClipboardData.GetDataPresent(DataFormats.Bitmap) Then

                Return mClipboardData.GetData(DataFormats.Bitmap)

            Else

                Return Nothing

            End If

        End Get

    End Property

    ''' <summary>
    '''   ウィンドウの修正率プロパティ
    ''' </summary>
    ''' <remarks>
    '''   クリップボード内に画像が存在する時はその画像の修正率を返す
    '''   クリップボード内に画像が存在しない時は修正率を１で返す（修正率は無しとする）
    ''' </remarks>
    Public ReadOnly Property WindowFixRate As Double

        Get

            'クリップボード内の画像が存在する時
            If Not ClipboardImage Is Nothing Then

                'クリップボード内の画像の修正率を返す ※修正率 = 画像の幅 / 画像の高さ
                Return Me.ClipboardImage.Width / Me.ClipboardImage.Height

            Else

                Return 1

            End If

        End Get

    End Property

    ''' <summary>
    '''   クリップボードに画像をコピーするか
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property IsCopyImageToClipboard As Boolean

        Get

            'クリップボードにコピーする画像パスが存在する時
            If Not _CopyImageToClipboardPath Is Nothing Then

                Return True

            Else

                Return False

            End If

        End Get

    End Property

    ''' <summary>
    '''   画像の自動保存を行うか
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property IsAutoSave As Boolean

        Get

            '画像の保存先パスが存在する時
            If Not String.IsNullOrEmpty(_SaveImagePath) Then

                Return True

            Else

                Return False

            End If

        End Get

    End Property

    ''' <summary>
    '''   ウィンドウを自動で閉じる時間プロパティ
    ''' </summary>
    ''' <remarks>
    '''   コマンドライン引数によってセットされる
    '''   セットされていない時はNothingを返す
    ''' </remarks>
    Public WriteOnly Property AutoCloseTime As TimeSpan

        Set(value As TimeSpan)

            _AutoCloseTimeSpan = value

        End Set

    End Property

    ''' <summary>
    '''   画像の保存先パスプロパティ
    ''' </summary>
    ''' <remarks>
    '''   コマンドライン引数によってセットされる
    '''   セットされていない時はNothingを返す
    ''' </remarks>
    Public WriteOnly Property SaveImagePath As String

        Set(value As String)

            _SaveImagePath = value

        End Set

    End Property

    ''' <summary>
    '''   画像の保存ファイル名
    ''' </summary>
    ''' <remarks></remarks>
    Public WriteOnly Property SaveImageFileName As String

        Set(value As String)

            _SaveImageFileName = value

        End Set

    End Property

    ''' <summary>
    '''   クリップボードにコピーする画像パスプロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public WriteOnly Property CopyImageToClipboardPath As String

        Set(value As String)

            _CopyImageToClipboardPath = value

        End Set

    End Property

    ''' <summary>
    '''   ウィンドウを非表示するかどうかプロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public Property DoNotShowWindow As Boolean

        Get

            Return _DoNotShowWindow

        End Get

        Set(value As Boolean)

            _DoNotShowWindow = value

        End Set

    End Property

    ''' <summary>
    '''   出力画像拡張子プロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public WriteOnly Property OutputImageExtension As OutputExtension

        Set(value As OutputExtension)

            _OutputImageExtension = value

        End Set

    End Property

    ''' <summary>
    '''   表示画像サイズプロパティ
    ''' </summary>
    ''' <remarks></remarks>
    Public WriteOnly Property DisplayImageSize As Double

        Set(value As Double)

            _DisplayImageSize = value

        End Set

    End Property

#End Region

#Region "DLL関数"

    ''' <summary>
    '''   指定されたウィンドウの表示状態、および通常表示のとき、最小化されたとき、最大化されたときの位置を返します。
    ''' </summary>
    ''' <param name="hWnd">ウィンドウのハンドル</param>
    ''' <param name="lpwndpl">位置データ</param>
    ''' <returns>
    '''   関数が成功すると、0 以外の値が返ります。
    '''   関数が失敗すると、0 が返ります。拡張エラー情報を取得するには、 関数を使います。
    ''' </returns>
    ''' <remarks>
    '''   この関数が取得する WINDOWPLACEMENT 構造体の flags メンバは、常に 0 です。
    '''   hWnd パラメータで指定したウィンドウが最大化されている場合、
    '''   showCmd メンバが SW_SHOWMAXIMIZED に設定されます。
    '''   ウィンドウが最小化されている場合は、showCmd メンバが SW_SHOWMINIMIZED に設定されます。
    '''   それ以外の場合は、SW_SHOWNORMAL に設定されます。
    '''   WINDOWPLACEMENT 構造体の length メンバは、sizeof(WINDOWPLACEMENT) に設定されていなければなりません。
    '''   このメンバが正しく設定されていないと、0（FALSE）が返ります。
    '''   ウィンドウの位置座標の正しい扱い方の詳細については、 構造体の説明を参照してください。
    ''' </remarks>
    <DllImport("user32.dll")> _
    Private Shared Function GetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    ''' <summary>
    '''   指定されたウィンドウの表示状態を設定し、そのウィンドウの通常表示のとき、最小化されたとき、および最大化されたときの位置を設定します
    ''' </summary>
    ''' <param name="hWnd">ウィンドウのハンドル</param>
    ''' <param name="lpwndpl">位置データ</param>
    ''' <returns>
    '''   関数が成功すると、0 以外の値が返ります。
    '''   関数が失敗すると、0 が返ります。拡張エラー情報を取得するには、 関数を使います。
    ''' </returns>
    ''' <remarks>
    '''   WINDOWPLACEMENT 構造体で指定された情報を適用するとウィンドウが完全に画面の外に出てしまう場合は、
    '''   ウィンドウが画面に現れるように座標が自動調整されます。
    '''   この調整では、画面の解像度の変更や複数モニタの構成も考慮されます。
    '''   WINDOWPLACEMENT 構造体の length メンバは、sizeof(WINDOWPLACEMENT) に設定されていなければなりません。
    '''   このメンバが正しく設定されていないと、0（FALSE）が返ります。
    '''   ウィンドウの位置座標の正しい扱い方の詳細については、 構造体の説明を参照してください。
    ''' </remarks>
    <DllImport("user32.dll")> _
    Private Shared Function SetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

#End Region

#Region "変数"

    ''' <summary>ウィンドウを閉じる用タイマー</summary>
    ''' <remarks></remarks>
    Private _CloseWindowTimer As System.Windows.Threading.DispatcherTimer

    ''' <summary>ウィンドウを自動で閉じる時間（プロパティで使用する）</summary>
    ''' <remarks></remarks>
    Private _AutoCloseTimeSpan As TimeSpan

    ''' <summary>画像の保存先パス（プロパティで使用する）</summary>
    ''' <remarks></remarks>
    Private _SaveImagePath As String

    ''' <summary>保存画像ファイル名（プロパティで使用する）</summary>
    ''' <remarks></remarks>
    Private _SaveImageFileName As String

    ''' <summary>クリップボードにコピーする画像パス（プロパティで使用する）</summary>
    ''' <remarks></remarks>
    Private _CopyImageToClipboardPath As String

    ''' <summary>ウィンドウを非表示するかどうか（プロパティで使用する）</summary>
    ''' <remarks></remarks>
    Private _DoNotShowWindow As Boolean = False

    ''' <summary>出力画像拡張子（プロパティで使用する）</summary>
    ''' <remarks></remarks>
    Private _OutputImageExtension As OutputExtension

    ''' <summary>表示画像サイズ（プロパティで使用する）</summary>
    ''' <remarks></remarks>
    Private _DisplayImageSize As Double

#End Region

#Region "イベント"

#Region "メインウィンドウ"

    ''' <summary>
    '''   MainWindowのLoadedイベント
    ''' </summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">Loadedイベント</param>
    ''' <remarks></remarks>
    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        '「クリップボードに画像をコピーする時」 かつ 「クリップボードに指定された画像をコピーが失敗した時」
        If Me.IsCopyImageToClipboard AndAlso Not _SaveImageToClipboard(_CopyImageToClipboardPath) Then

            'ウィンドウを閉じる
            Me.Close()

        End If

        'クリップボードに画像が存在しない時はLoadedイベントを終了
        If Me.ClipboardImage Is Nothing Then Exit Sub

        'メインウィンドウ設定
        _ToSetupMainWindow()

        'クリップボード内画像表示コントロール設定
        _ToSetupImgClipboardImage()

        '画像の自動保存を行う時
        If Me.IsAutoSave Then

            '保存ファイル名を作成（フルパス）
            Dim mSaveImageFile As String = _SaveImagePath & _SaveImageFileName & "." & _OutputImageExtension.ToString

            'クリップボード画像の保存処理
            _SaveClipboardImage(mSaveImageFile, _OutputImageExtension)

        End If

    End Sub

    ''' <summary>
    '''   MainWindowのContentRenderedイベント
    ''' </summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">ContentRenderedイベント</param>
    ''' <remarks>ウィンドウのコンテンツが描画された後に発生</remarks>
    Private Sub MainWindow_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered

        'クリップボードに画像が存在しない時はウィンドウを閉じる
        If Me.ClipboardImage Is Nothing Then MainWindow_Unloaded(Me, Nothing)

    End Sub

    ''' <summary>
    '''   MainWindowのStateChangedイベント
    ''' </summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">StateChangedイベント</param>
    ''' <remarks>最小化・最大化を無効化する（一瞬最大化・最小化処理がされてしまう……）</remarks>
    Private Sub MainWindow_StateChanged(sender As Object, e As EventArgs) Handles Me.StateChanged

        'ウィンドウの状態が「最小化」 または ウィンドウの状態が「最大化」の時
        If Me.WindowState = WindowState.Minimized OrElse Me.WindowState = Windows.WindowState.Maximized Then

            'ウィンドウの状態を「通常にする」
            Me.WindowState = Windows.WindowState.Normal

        End If

    End Sub

    ''' <summary>
    '''   MainWindowのMouseLeftButtonDownイベント
    ''' </summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">MouseButtonEventArgs</param>
    ''' <remarks>MainWindow上にあるとき（またはマウスがキャプチャされたとき）にマウスの左ボタンが押されると発生する</remarks>
    Private Sub MainWindow_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonDown

        '左クリック押下でドラッグしウインドウを移動出来るようにする
        Me.DragMove()

    End Sub

    ''' <summary>
    '''   MainWindowのUnloadedイベント
    ''' </summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">Unloadedイベント</param>
    ''' <remarks></remarks>
    Private Sub MainWindow_Unloaded(sender As Object, e As RoutedEventArgs) Handles Me.Unloaded

        'オーナーウィンドウが存在する時
        If Not Me.Owner Is Nothing Then

            'オーナーウィンドウを閉じる
            '※Me.Close()だけだとプログラムが終了しないのでここでオーナーウィンドウを閉じるようにする
            Me.Owner.Close()

        Else

            '自分自身を閉じる
            Me.Close()

        End If

    End Sub

#End Region

#Region "コンテキストメニュー"

    ''' <summary>
    '''   コンテキストメニューClickイベント
    ''' </summary>
    ''' <param name="sender">MenuItem</param>
    ''' <param name="e">コンテキストメニューClickイベント</param>
    ''' <remarks></remarks>
    Private Sub ContextMenuItem_Click(sender As Object, e As System.EventArgs)

        'クリックされたアイテム名を取得する
        Dim mClickItemName As String = DirectCast(sender, System.Windows.Controls.MenuItem).Header

        Select Case mClickItemName

            Case ContextMenuItem.クリップボードの画像を保存.ToString

                '名前を付けて保存ダイアログを表示
                Dim mDialog As System.Windows.Forms.SaveFileDialog = _GetSaveAsDialog(_SaveImageFileName)

                '名前をつけて保存ダイアログでOKが押されたら
                If mDialog.ShowDialog = Windows.Forms.DialogResult.OK Then

                    'ファイル名から拡張子文字列を取得 ※「.拡張子」で取得されるため、「.」以降の文字列を取得
                    Dim mExtensionString As String = System.IO.Path.GetExtension(mDialog.FileName).Substring(1)

                    '「画像の出力形式」列挙体に変換が出来なかった時
                    Dim mExtension As OutputExtension
                    If Not [Enum].TryParse(Of OutputExtension)(mExtensionString.ToLower, mExtension) Then

                        '「ファイルの種類に存在しない拡張子を指定して画像の保存はできません」メッセージを表示
                        MessageBox.Show(_cMessage.NotExistsExtensionForSaveAsDialog _
                                      , cNameSpaceName _
                                      , MessageBoxButton.OK _
                                      , MessageBoxImage.Error)

                        '処理を終了
                        Exit Sub

                    End If

                    'クリップボード内の画像の保存処理
                    _SaveClipboardImage(mDialog.FileName, mExtension)

                End If

            Case ContextMenuItem.閉じる.ToString

                Me.Close()

        End Select

    End Sub

#End Region

#Region "その他"

    ''' <summary>
    '''   SourceInitializedイベントを発生
    ''' </summary>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>ウインドウの初期化中に呼び出されます</remarks>
    Protected Overrides Sub OnSourceInitialized(ByVal e As EventArgs)

        '基底クラスのSourceInitializedイベントを発生させる
        MyBase.OnSourceInitialized(e)

        'WPFコンテンツを格納するWin32のウィンドウを取得する
        Dim mHwndSource As HwndSource = CType(HwndSource.FromVisual(Me), HwndSource)

        'ウィンドウメッセージを受信するイベントハンドラーを追加
        mHwndSource.AddHook(AddressOf WndHookProc)

    End Sub

    ''' <summary>MainWindowのClosingイベント</summary>
    ''' <param name="e">キャンセルできるイベントのデータ</param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnClosing(ByVal e As CancelEventArgs)

        '----------------------------------------------
        ' ウィンドウプロシージャをフックする設定を解除
        '----------------------------------------------
        'WPFコンテンツを格納するWin32のウィンドウを取得する
        Dim mHwndSource As HwndSource = DirectCast(PresentationSource.FromVisual(Me), HwndSource)

        'Win32のウィンドウが存在する時、ウィンドウメッセージを受信するイベントハンドラーを削除
        If Not mHwndSource Is Nothing Then mHwndSource.RemoveHook(AddressOf WndHookProc)

        '----------------------------------------------
        ' Closedイベントを発生
        '----------------------------------------------
        '基底クラスのClosedイベントを発生させる
        MyBase.OnClosing(e)

    End Sub

    ''' <summary>
    '''   ウィンドウプロシージャをフック
    ''' </summary>
    ''' <param name="hwnd">ウィンドウのハンドル</param>
    ''' <param name="msg">メッセージの識別子</param>
    ''' <param name="wParam">メッセージの最初のパラメータ</param>
    ''' <param name="lParam">メッセージの２番目のパラメータ</param>
    ''' <param name="handled">ハンドルフラグ</param>
    ''' <returns>0 に初期化されたポインターまたはハンドル</returns>
    ''' <remarks>
    '''   ウインドウプロシージャ
    '''     メッセージを処理する専用のルーチン
    '''   Hook（フック）
    '''     独自の処理を割り込ませるための仕組み
    '''      注意：デバッグ時はこのメソッドの処理で止まりません。処理を確認した時は「System.Diagnostics.DebuggerStepThrough()」行を削除して下さい
    ''' </remarks>
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Function WndHookProc(ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByRef handled As Boolean) As IntPtr

        If msg = cWM_SIZING.Message Then

            'アスペクト比を保ったままウィンドウサイズを変更
            Call _ResizeWindowKeepingAspectRatio(wParam, lParam)

        ElseIf msg = cWM_NCHITTEST.Message Then

            '現在のマウス位置を返す
            Return _GetMousePotisionInTheForm(lParam, handled)

        End If

        Return IntPtr.Zero

    End Function

#End Region

#End Region

#Region "メソッド"

#Region "ウィンドウプロシージャをフック関連"

    ''' <summary>
    '''   ウィンドウサイズをアスペクト比率を保って変更
    ''' </summary>
    ''' <param name="wParam">メッセージの最初のパラメータ</param>
    ''' <param name="lParam">メッセージの２番目のパラメータ</param>
    ''' <remarks>
    '''   「ウィンドウプロシージャをフック」するイベントから呼ばれます
    '''    ※ウィンドウのサイズを変更した時、アスペクト比を保ったままサイズが変更されるようにするためのメソッド
    ''' </remarks>
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub _ResizeWindowKeepingAspectRatio(ByVal wParam As IntPtr, ByVal lParam As IntPtr)

        'アンマネージメモリのRECT構造体をマネージオブジェクト（RECT構造体）にデータをマーシャリングする
        '※ウィンドウプロシージャに渡ってきた「lParam」を.NET側で使えるようにデータを変換する。Marshalingは「整列」という意味の英単語
        Dim mRect As RECT = Marshal.PtrToStructure(lParam, GetType(RECT))

        'ウィンドウの幅と高さを求める
        Dim mWindowWidth As Double = mRect.Right - mRect.Left
        Dim mWindowHeight As Double = mRect.Bottom - mRect.Top

        'ウィンドウの幅と高さの増減値を取得  
        ' ウィンドウ幅  の増減値：「(ウィンドウ高さ * 修正率) - ウィンドウ幅  」
        ' ウィンドウ高さの増減値：「(ウィンドウ幅   * 修正率) - ウィンドウ高さ」
        Dim mChangeWidth As Double = Math.Round((mWindowHeight * Me.WindowFixRate)) - mWindowWidth
        Dim mChangeHeight As Double = Math.Round((mWindowWidth / Me.WindowFixRate)) - mWindowHeight

        Select Case wParam.ToInt32()

            Case cWM_SIZING.wParam.WMSZ_LEFT, cWM_SIZING.wParam.WMSZ_RIGHT

                '「左端」と「右端」の時は、ウインドウ幅の増減値を右下隅のＹ座標に設定
                mRect.Bottom = mRect.Bottom + mChangeHeight

            Case cWM_SIZING.wParam.WMSZ_TOP, cWM_SIZING.wParam.WMSZ_BOTTOM

                '「上端」と「下端」の時は、ウインドウ高さの増減値を右下隅のＸ座標に設定
                mRect.Right = mRect.Right + mChangeWidth

            Case cWM_SIZING.wParam.WMSZ_TOPLEFT

                'ウィンドウ幅の増減値が０より大きい時
                If (mChangeWidth > 0) Then

                    'ウィンドウの左位置を再設定「ウィンドウの左位置 - ウィンドウ幅の増減値」
                    mRect.Left = mRect.Left - mChangeWidth

                Else

                    'ウィンドウの上位置を再設定「ウィンドウの上位置 - ウィンドウ高さの増減値」
                    mRect.Top = mRect.Top - mChangeHeight

                End If

            Case cWM_SIZING.wParam.WMSZ_TOPRIGHT

                'ウィンドウ幅の増減値が０より大きい時
                If (mChangeWidth > 0) Then

                    'ウィンドウの右位置を再設定「ウィンドウの右位置 + ウィンドウ幅の増減値」
                    mRect.Right = mRect.Right + mChangeWidth

                Else

                    'ウィンドウの上位置を再設定「ウィンドウの上位置 - ウィンドウ高さの増減値」
                    mRect.Top = mRect.Top - mChangeHeight

                End If

            Case cWM_SIZING.wParam.WMSZ_BOTTOMLEFT

                'ウィンドウ幅の増減値が０より大きい時
                If (mChangeWidth > 0) Then

                    'ウィンドウの左位置を再設定「ウィンドウの左位置 - ウィンドウ幅の増減値」
                    mRect.Left = mRect.Left - mChangeWidth

                Else

                    'ウィンドウの下位置を再設定「ウィンドウの下位置 + ウィンドウ高さの増減値」
                    mRect.Bottom = mRect.Bottom + mChangeHeight

                End If

            Case cWM_SIZING.wParam.WMSZ_BOTTOMRIGHT

                'ウィンドウ幅の増減値が０より大きい時
                If (mChangeWidth > 0) Then

                    'ウィンドウの右位置を再設定「ウィンドウの右位置 + ウィンドウ幅の増減値」
                    mRect.Right = mRect.Right + mChangeWidth

                Else

                    'ウィンドウの下位置を再設定「ウィンドウの下位置 + ウィンドウ高さの増減値」
                    mRect.Bottom = mRect.Bottom + mChangeHeight

                End If

        End Select

        'マネージオブジェクト（RECT構造体）をアンマネージメモリブロックにデータをマーシャリングする
        '※この処理で変更したRECT構造体の値を
        Marshal.StructureToPtr(mRect, lParam, False)

    End Sub

    ''' <summary>
    '''   マウス位置を取得
    ''' </summary>
    ''' <param name="lParam">メッセージの２番目のパラメータ</param>
    ''' <param name="handled">ハンドルフラグ</param>
    ''' <returns>現在のマウス位置を返す</returns>
    ''' <remarks>
    '''  ・「ウィンドウプロシージャをフック」するイベントから呼ばれます
    '''  ・現在のウィンドウのスタイルだと右下の部分でしかウィンドウサイズの変更が不可能である
    '''      スタイルの設定：WindowStyle="None" AllowsTransparency="True" ResizeMode="CanResizeWithGrip"
    '''    そのためウィンドウの端に来た時、リサイズ可能領域内（マウス位置がキャプションバー内）であることを知らせる
    '''  ・スクリーン座標とクライアント座標について
    '''      スクリーン座標            ：画面の左上隅の点を原点とした座標
    '''      フォームのクライアント座標：フォームの描画可能なクライアント領域の左上隅の点を原点とした座標
    '''      ※参考URL：http://www.atmarkit.co.jp/fdotnet/dotnettips/377screentoclient/screentoclient.html
    ''' </remarks>
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Function _GetMousePotisionInTheForm(ByVal lParam As IntPtr, ByRef handled As Boolean) As IntPtr

        'これ以上処理させない（完全に処理を横取りする）
        handled = True

        '------------------------------------
        ' クライアント座標のマウス位置を取得
        '------------------------------------
        'スクリーン座標のマウス位置を取得
        Dim mMousePositionOnScreen As New System.Windows.Point(CInt(lParam) And &HFFFF, (CInt(lParam) >> 16) And &HFFFF)

        'スクリーン座標のマウス位置をクライアント座標のマウス位置に変換
        Dim mMousePositionOnClient As System.Windows.Point = PointFromScreen(mMousePositionOnScreen)

        '------------------------------------
        ' リサイズ可能とするサイズを取得
        '------------------------------------
        'ウィンドウの周囲にある水平サイズ変更境界の高さサイズを取得
        Dim ResizableHorizontal As Double = SystemParameters.ResizeFrameHorizontalBorderHeight

        'ウィンドウの周囲にある垂直サイズ変更境界の幅サイズを取得
        Dim ResizableVertical As Double = SystemParameters.ResizeFrameVerticalBorderWidth

        'タイトルバーの高さを取得
        Dim ResizableCaptionHeader As Double = SystemParameters.CaptionHeight

        '------------------------------------
        ' 四隅の斜め方向にリサイズ可能
        '------------------------------------
        '左上の斜め方向にリサイズ可能
        If New System.Windows.Rect(0, 0, ResizableVertical, ResizableHorizontal).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTTOPLEFT)

        '右上の斜め方向にリサイズ可能
        If New System.Windows.Rect(Me.Width - ResizableVertical, 0, ResizableVertical, ResizableHorizontal).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTTOPRIGHT)

        '左下の斜め方向にリサイズ可能
        If New System.Windows.Rect(0, Height - ResizableHorizontal, ResizableVertical, ResizableHorizontal).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTBOTTOMLEFT)

        '右下の斜め方向にリサイズ可能
        If New System.Windows.Rect(Me.Width - ResizableVertical, Me.Height - ResizableHorizontal, ResizableVertical, ResizableHorizontal).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTBOTTOMRIGHT)

        '------------------------------------
        ' 四辺の直交方向にリサイズ可能
        '------------------------------------
        '上に直交方向にリサイズ可能
        If New System.Windows.Rect(0, 0, Me.Width, ResizableVertical).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTTOP)

        '左に直交方向にリサイズ可能
        If New System.Windows.Rect(0, 0, ResizableVertical, Me.Height).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTLEFT)

        '右に直交方向にリサイズ可能
        If New System.Windows.Rect(Me.Width - ResizableVertical, 0, ResizableVertical, Me.Height).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTRIGHT)

        '下に直交方向にリサイズ可能
        If New System.Windows.Rect(0, Me.Height - ResizableHorizontal, Me.Width, ResizableHorizontal).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTBOTTOM)

        '------------------------------------
        ' タイトルバーにマウスがあるか判断
        '------------------------------------
        'マウスがタイトルバーにある
        If New System.Windows.Rect(0, 0, Me.Width, ResizableCaptionHeader).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTCAPTION)

        '上記以外はクライアント領域とする
        Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTCLIENT)

    End Function

#End Region

#Region "コントロール設定関連関連"

    ''' <summary>
    '''   メインウィンドウ設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub _ToSetupMainWindow()

        '----------------------------------
        ' ウィンドウの表示設定
        '----------------------------------
        'ウィンドウを非表示するかどうかがTrueの時はウィンドウを非表示
        If _DoNotShowWindow = True Then Me.Visibility = Windows.Visibility.Collapsed

        'ウィンドウを透明化 ※透明にしておかないと、ウィンドウの移動している処理が見えてしまう……
        Me.Opacity = 0

        '----------------------------------
        ' ウィンドウの最大・最小サイズ設定
        '----------------------------------
        'ウィンドウの最大サイズ設定 ※画像の幅、高さを設定
        Me.MaxWidth = Me.ClipboardImage.Width
        Me.MaxHeight = Me.ClipboardImage.Height

        'ウィンドウの最小サイズ設定 ※画像サイズの１％のサイズ
        Me.MinWidth = Me.ClipboardImage.Width * Math.Sqrt(0.01)
        Me.MinHeight = Me.ClipboardImage.Height * Math.Sqrt(0.01)

        '----------------------------------
        ' ウィンドウの表示サイズ設定
        '----------------------------------
        'マウスが存在するディスプレイの４分の１のサイズを取得する
        Dim mQuarterScreenSize As System.Drawing.Size = _GetScreenSizeExistsMouseOnScreen(0.25)

        'ウィンドウの表示画像サイズ％を取得
        Dim mDisplayWindowWidth As Double = Me.MaxWidth * Math.Sqrt(_DisplayImageSize * 0.01)
        Dim mDisplayWindowHeight As Double = Me.MaxHeight * Math.Sqrt(_DisplayImageSize * 0.01)

        '       ディスプレイの４分の１のサイズの幅より表示するウィンドウの幅が大きい
        'または ディスプレイの４分の１のサイズの高さより表示するウィンドウの高さが大きい時
        If mQuarterScreenSize.Width < mDisplayWindowWidth OrElse mQuarterScreenSize.Height < mDisplayWindowHeight Then

            'ウィンドウのサイズを表示画像サイズ％に設定
            Me.Width = mDisplayWindowWidth
            Me.Height = mDisplayWindowHeight

        Else

            'ウィンドウのサイズをデフォルトサイズに設定
            Me.Width = Me.MaxWidth
            Me.Height = Me.MaxHeight

        End If

        '----------------------------------
        ' ウィンドウの表示位置設定
        '----------------------------------
        'ウィンドウの表示位置を取得 ※マウスの右下に表示
        Dim mMousePosition = _GetMousePosition()
        Dim mWindowLeft As Double = mMousePosition.X + _cDisplayWindowPlaceMargin
        Dim mWindowTop As Double = mMousePosition.Y + _cDisplayWindowPlaceMargin

        'マウスが存在するディスプレイのサイズを取得する
        Dim mScreenSize As System.Drawing.Size = _GetScreenSizeExistsMouseOnScreen()

        '※表示するウィンドウの下の位置がマウスが存在するディスプレイ外にある時はウィンドウの上の位置を修正する
        '「ディスプレイの下位置 - ウィンドウの上位置（ディスプレイ内に表示されているウィンドウの縦サイズ）」を計算
        Dim mWindowVerticalSizeOnScreen As Double = mScreenSize.Height - mWindowTop

        'ディスプレイ内に表示されているウィンドウの縦サイズがウィンドウの高さより小さい時
        If mWindowVerticalSizeOnScreen < Me.Height Then

            'ディスプレイ外に表示されているウィンドウ縦サイズを計算
            Dim mWindowVerticalSizeOutScreen As Double = Me.Height - mWindowVerticalSizeOnScreen

            'ウィンドウの表示上位置をディスプレイ外に表示されているウィンドウ縦サイズ分上に修正
            Me.Top = mWindowTop - mWindowVerticalSizeOutScreen

        Else

            'ウィンドウの通常の表示位置へ移動
            Me.Top = mWindowTop

        End If

        '※表示するウィンドウの右の位置が仮想画面の左側外にある時はウィンドウの左位置を修正する
        '仮想画面の右側、ウィンドウの右位置を計算
        Dim mVirtualScreenRight As Double = SystemParameters.VirtualScreenLeft + SystemParameters.VirtualScreenWidth
        Dim mWindowRight As Double = mWindowLeft + Me.Width

        '仮想画面の右側よりウィンドウの右位置が大きい時
        If mVirtualScreenRight < mWindowRight Then

            Me.Left = mWindowLeft - (mWindowRight - mVirtualScreenRight)

        Else

            Me.Left = mWindowLeft

        End If

        'ウィンドウを透明を解除
        Me.Opacity = 100

        '----------------------------------
        ' 「Alt+Tab」ウィンドウから非表示
        '----------------------------------
        '「Alt+Tab」ウィンドウから非表示にする
        '※この処理を先に行うと、ウィンドウが表示されてしまうため、後で処理する
        _HideAltTabWindow()

        '----------------------------------
        ' ウィンドウの閉じる時間設定
        '----------------------------------
        'ウィンドウを閉じる時間をセット
        If Not _AutoCloseTimeSpan = Nothing Then _ToSetupCloseWindowTimer(_AutoCloseTimeSpan)

    End Sub

    ''' <summary>
    '''   「Alt+Tab」ウインドウに表示させない
    ''' </summary>
    ''' <remarks>
    '''   Alt+Tab ダイアログに表示されない Window をオーナーに設定し
    '''   ShowInTaskbar プロパティに False を設定することでAlt+Tabに表
    '''   示されなくなる。オーナーウィンドウを非表示にしておけば、
    '''   SingleBorderWindow や ThreeDBorderWindow の Window が 
    '''   Alt+Tab ダイアログに表示されないようにすることができます。
    ''' </remarks>
    Private Sub _HideAltTabWindow()

        'オーナー用ウィンドウを作成
        Dim mOwnerWindow As New Window

        'Alt＋Tabに表示されない設定にする
        mOwnerWindow.WindowStyle = Windows.WindowStyle.ToolWindow
        mOwnerWindow.ShowInTaskbar = False

        '表示領域外にオーナー用ウィンドウを表示させる（一旦表示させる必要があるため）
        mOwnerWindow.Left = -100
        mOwnerWindow.Height = 0
        mOwnerWindow.Width = 0
        mOwnerWindow.Show()

        'オーナー用ウィンドウを非表示
        mOwnerWindow.Hide()

        'オーナーウィンドウにオーナー用ウィンドウを設定する
        Me.Owner = mOwnerWindow

    End Sub

    ''' <summary>
    '''   ウィンドウを閉じる用タイマー設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub _ToSetupCloseWindowTimer(ByVal pCloseTime As TimeSpan)

        'インスタンスを作成
        _CloseWindowTimer = New System.Windows.Threading.DispatcherTimer

        'ウィンドウを閉じる用タイマーにTickイベントを設定
        '※イベントの中の処理が１行で済む場合はラムダ式で記述したほうがスッキリする
        AddHandler _CloseWindowTimer.Tick, Sub(sender As Object, e As EventArgs) Me.Close()

        'Tickイベントが発生する間隔を設定
        _CloseWindowTimer.Interval = pCloseTime

        'ウィンドウを閉じる用タイマーを起動
        _CloseWindowTimer.Start()

    End Sub

    ''' <summary>
    '''   クリップボード内画像表示コントロール設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub _ToSetupImgClipboardImage()

        'クリップボードの画像をセット
        imgClipboardImage.Source = Me.ClipboardImage

        '右クリックメニューをセット
        imgClipboardImage.ContextMenu = _CreateContextMenu()

    End Sub

    ''' <summary>
    '''   コンテキストメニュー作成
    ''' </summary>
    ''' <returns>コンテキストメニュー</returns>
    ''' <remarks></remarks>
    Private Function _CreateContextMenu() As ContextMenu

        Dim mContextMenu As New System.Windows.Controls.ContextMenu

        'コンテキストメニューアイテム列挙体数分繰り返す
        For Each mItemName As String In System.Enum.GetNames(GetType(ContextMenuItem))

            Dim mAddMenuItem As New System.Windows.Controls.MenuItem

            'メニュー名とクリック時のイベントを設定
            mAddMenuItem.Header = mItemName
            AddHandler DirectCast(mAddMenuItem, System.Windows.Controls.MenuItem).Click, New System.Windows.RoutedEventHandler(AddressOf ContextMenuItem_Click)

            'コンテキストメニューにアイテムを追加
            mContextMenu.Items.Add(mAddMenuItem)

        Next

        Return mContextMenu

    End Function

#End Region

#Region "クリップボード関連"

    ''' <summary>
    '''   クリップボード内の画像の保存処理
    ''' </summary>
    ''' <param name="pSaveFileName">保存ファイル名（フルパス）</param>
    ''' <param name="pExtension">保存ファイル拡張子</param>
    ''' <remarks></remarks>
    Private Sub _SaveClipboardImage(ByVal pSaveFileName As String, ByVal pExtension As OutputExtension)

        Try

            Using mStream As System.IO.FileStream = New System.IO.FileStream(pSaveFileName, IO.FileMode.Create)

                Dim mEncoder As Object

                '拡張子によりエンコーダーのインスタンスを作成
                Select Case pExtension

                    Case OutputExtension.bmp

                        mEncoder = New BmpBitmapEncoder

                    Case OutputExtension.png

                        mEncoder = New PngBitmapEncoder

                    Case OutputExtension.gif

                        mEncoder = New GifBitmapEncoder

                    Case OutputExtension.jpg

                        mEncoder = New JpegBitmapEncoder

                    Case Else

                        mEncoder = New BmpBitmapEncoder

                End Select

                'クリップボード内の画像の保存処理
                mEncoder.Frames.Add(BitmapFrame.Create(CType(Me.ClipboardImage, BitmapSource)))
                mEncoder.Save(mStream)

            End Using

        Catch ex As Exception

            '「画像の保存に失敗しました」メッセージを表示
            MessageBox.Show(_cMessage.SaveImageFailure _
                          , cNameSpaceName _
                          , MessageBoxButton.OK _
                          , MessageBoxImage.Error)

            'エラーの内容をイベントログに書き込む
            Application.WriteEventLog(ex, _cMessage.SaveImageFailure, EventLogEntryType.Error)

        End Try

    End Sub

    ''' <summary>
    '''   クリップボードへ対象パスの画像をコピーする
    ''' </summary>
    ''' <param name="pPath">画像のパス（フルパス）</param>
    ''' <returns>True：画像のコピー成功、False：画像のコピー失敗</returns>
    ''' <remarks>
    '''   ※対応ファイル形式：BMP、GIF、EXIF、JPG、PNG、TIFF、ICO
    '''     System.Windows.Clipboard.SetData(DataFormats.Bitmap, mBitmap)だとエラーが
    '''     出てしまうため、System.Windows.Forms.Clipboardを使用する
    '''     ToDo:対応ファイル形式については調べる必要あり
    ''' </remarks>
    Private Function _SaveImageToClipboard(ByVal pPath As String) As Boolean

        Try

            'パスから画像データを取得 ※対応ファイル形式：BMP、GIF、EXIF、JPG、PNG、TIFF、ICO
            Dim mBitmap As New System.Drawing.Bitmap(pPath)

            '画像データをクリップボードにコピーする
            System.Windows.Forms.Clipboard.SetImage(mBitmap)

            '画像データを破棄する
            mBitmap.Dispose()

        Catch ex As Exception

            '「クリップボードへ画像のコピーが失敗しました」メッセージを表示
            MessageBox.Show(_cMessage.CopyImageToClipboardFailure _
                          , cNameSpaceName _
                          , MessageBoxButton.OK _
                          , MessageBoxImage.Error)

            'エラーの内容をイベントログに書き込む
            Application.WriteEventLog(ex, Application.GetErrorMessage(ex, Application.ErrorMessageType.EventLog), EventLogEntryType.Error)

            Return False

        End Try

        Return True

    End Function

#End Region

#Region "その他"

    ''' <summary>
    '''   マウス座標を取得
    ''' </summary>
    ''' <returns>マウス座標</returns>
    ''' <remarks></remarks>
    Private Function _GetMousePosition() As System.Windows.Point

        'マウス座標を取得 ※物理座標で取得される
        Dim mMousePosition As System.Drawing.Point = System.Windows.Forms.Cursor.Position
        Dim mMousePositionForWPF As New System.Windows.Point(mMousePosition.X, mMousePosition.Y)

        'マウス座標から論理座標へ変換
        Dim mSrc As PresentationSource = PresentationSource.FromVisual(Me)
        Dim mCt As CompositionTarget = mSrc.CompositionTarget
        Dim mLogicalPosition As System.Windows.Point = mCt.TransformFromDevice.Transform(mMousePositionForWPF)

        Return mLogicalPosition

    End Function

    ''' <summary>
    '''   マウスが存在しているディスプレイの任意のサイズを取得
    ''' </summary>
    ''' <param name="pSize">サイズ</param>
    ''' <returns>ディスプレイサイズ</returns>
    ''' <remarks>
    '''   ・pSizeに「0.25」を指定した場合、マウスが存在しているディスプレイの４分の１のサイスを取得出来る
    '''   ・何も指定しなかった場合はディスプレイのサイズを取得する
    '''   ※WPFアプリケーションにはウィンドウが存在しているディスプレイのサイズを取得する方法が存在しない
    '''     ためウィンドウズフォームアプリケーションのWindowクラスを使用してディスプレイサイズを取得する
    ''' </remarks>
    Private Function _GetScreenSizeExistsMouseOnScreen(Optional pSize As Double = 1) As System.Drawing.Size

        'ウィンドウズフォームを透明状態で表示
        Dim mForm As New System.Windows.Forms.Form
        With mForm
            .ShowInTaskbar = False
            .Opacity = 0
        End With
        mForm.Show()

        'マウスの位置を取得しウィンドウをマウスの位置に移動
        Dim mMousePosition As System.Drawing.Point = System.Windows.Forms.Cursor.Position
        mForm.Location = mMousePosition

        'ウィンドウが存在しているディスプレイのサイズを取得
        Dim mScreenSizeOnWindow As System.Windows.Forms.Screen = System.Windows.Forms.Screen.FromControl(mForm)

        'ウィンドウを閉じる
        mForm.Close()

        'ディスプレイの任意のサイズを取得
        Dim mScreenWidth As Integer = mScreenSizeOnWindow.Bounds.Width * Math.Sqrt(pSize)
        Dim mScreenHeight As Integer = mScreenSizeOnWindow.Bounds.Height * Math.Sqrt(pSize)

        Return New System.Drawing.Size(mScreenWidth, mScreenHeight)

    End Function

    ''' <summary>
    '''   名前を付けて保存ダイアログを取得
    ''' </summary>
    ''' <param name="pSaveFileName">保存ファイル名</param>
    ''' <param name="pDirectory">起動ディレクトリ</param>
    ''' <returns>ファイル形式にあった設定の名前を付けて保存ダイアログ</returns>
    ''' <remarks>
    '''   ファイル形式用の名前を付けて保存ダイアログを作成し返す
    ''' </remarks>
    Private Function _GetSaveAsDialog(ByVal pSaveFileName As String, Optional ByVal pDirectory As String = "") As System.Windows.Forms.SaveFileDialog

        '名前を付けて保存ダイアログを表示
        Dim mDailog As New System.Windows.Forms.SaveFileDialog
        With mDailog

            'デフォルトファイル設定（保存ファイル名）
            .FileName = pSaveFileName

            '表示ファイル設定
            .Filter = cSaveDialogExtensionFilter()

            '起動ディレクトリ設定
            .InitialDirectory = pDirectory

        End With

        Return mDailog

    End Function

#End Region

#End Region

End Class