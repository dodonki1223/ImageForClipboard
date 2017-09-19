Option Explicit On

Imports System.Runtime.InteropServices

''' <summary>
'''   ImageForClipboardで使用する定義情報を提供する
''' </summary>
''' <remarks></remarks>
Public Class ImageForClipboardDefinition

#Region "定数"

    ''' <summary>名前空間名</summary>
    ''' <remarks>「ImageForClipboard」が返ります</remarks>
    Public Shared cNameSpaceName As String = GetType(ImageForClipboard.Application).Namespace.ToString

    ''' <summary>画像のデフォルト保存ファイル名</summary>
    ''' <remarks></remarks>
    Public Const cDefaultSaveImageFileName As String = "クリップボードの画像"

    ''' <summary>名前を付けて保存ダイアログのファイルの種類フィルター文字列</summary>
    ''' <remarks></remarks>
    Public Shared Function cSaveDialogExtensionFilter() As String

        Dim mFilterString As New System.Text.StringBuilder

        '画像の出力形式列挙体分繰り返す
        For Each mExtension As OutputExtension In [Enum].GetValues(GetType(OutputExtension))

            mFilterString.Append(mExtension.ToString.ToUpper & "ファイル|" & "*." & mExtension.ToString.ToLower & "|")

        Next

        mFilterString.Append("すべてのファイル|*.*")

        Return mFilterString.ToString

    End Function

#Region "ウィンドウのサイズ変更関連定数"

    ''' <summary>WM_SIZING Message定数群</summary>
    ''' <remarks></remarks>
    Public Class cWM_SIZING

        ''' <summary>Message番号</summary>
        ''' <remarks>
        '''   ユーザーがサイズ変更中のウィンドウに送信されます。
        '''   このメッセージを処理することにより、アプリケーションはドラッグ矩形のサイズと位置を監視し、
        '''   必要に応じてそのサイズや位置を変更することができます。
        '''   ウィンドウは、このメッセージをWindowProc関数を通じて受信します。
        '''   参考URL：https://msdn.microsoft.com/ja-jp/library/windows/desktop/ms632647(v=vs.85).aspx
        ''' </remarks>
        Public Const Message As Integer = &H214

        ''' <summary>サイズ変更後の位置を取得するパラメータ</summary>
        ''' <remarks></remarks>
        Public Enum wParam

            ''' <summary>左端</summary>
            WMSZ_LEFT = 1

            ''' <summary>右端</summary>
            WMSZ_RIGHT = 2

            ''' <summary>上端</summary>
            WMSZ_TOP = 3

            ''' <summary>左上隅</summary>
            WMSZ_TOPLEFT = 4

            ''' <summary>右上隅</summary>
            WMSZ_TOPRIGHT = 5

            ''' <summary>下端</summary>
            WMSZ_BOTTOM = 6

            ''' <summary>左下隅</summary>
            WMSZ_BOTTOMLEFT = 7

            ''' <summary>右下隅</summary>
            WMSZ_BOTTOMRIGHT = 8

        End Enum

    End Class

    ''' <summary>WM_NCHITTEST Message定数群</summary>
    ''' <remarks></remarks>
    Public Class cWM_NCHITTEST

        ''' <summary>Message番号</summary>
        ''' <remarks>
        '''   特定の画面座標に対応するウィンドウの部分を決定するためにウィンドウに送信されます。
        '''   これは、たとえば、カーソルが移動したとき、マウスボタンが押されたとき、
        '''   または解放されたとき、またはWindowFromPointなどの関数の呼び出しに応答して起こることがあります。
        '''   マウスがキャプチャされていない場合、メッセージはカーソルの下のウィンドウに送信されます。
        '''   それ以外の場合、メッセージはマウスをキャプチャしたウィンドウに送信されます。
        '''   ウィンドウは、このメッセージをWindowProc関数を通じて受信します。
        '''   参考URL：https://msdn.microsoft.com/ja-jp/library/windows/desktop/ms645618(v=vs.85).aspx
        ''' </remarks>
        Public Const Message As Integer = &H84

        ''' <summary>カーソルのホットスポットの位置</summary>
        ''' <remarks></remarks>
        Public Enum CursorHotSpot

            ''' <summary>サイズ変更の境界線を持たないウィンドウの境界線です</summary>
            HTBORDER = 18

            ''' <summary>サイズ変更可能なウィンドウの下水平境界線で（ユーザーはマウスをクリックしてウィンドウを垂直方向にサイズ変更できます）</summary>
            HTBOTTOM = 15

            ''' <summary>サイズ変更可能なウィンドウの境界線の左下隅（マウスをクリックすると、ウィンドウのサイズを斜めに変更できます）</summary>
            HTBOTTOMLEFT = 16

            ''' <summary>サイズ変更可能なウィンドウの境界線の右下にあります（マウスをクリックすると、ウィンドウのサイズを斜めに変更できます）</summary>
            HTBOTTOMRIGHT = 17

            ''' <summary>タイトルバー</summary>
            HTCAPTION = 2

            ''' <summary>クライアントエリア</summary>
            HTCLIENT = 1

            ''' <summary>閉じるボタン</summary>
            HTCLOSE = 20

            ''' <summary>画面の背景またはウィンドウ間の分割線（HTNOWHEREと同じですが、DefWindowProc関数でエラーを示すビープ音が鳴ります</summary>
            HTERROR = (-2)

            ''' <summary>サイズボックス（HTSIZEと同じ）</summary>
            HTGROWBOX = 4

            ''' <summary>ヘルプボタン</summary>
            HTHELP = 21

            ''' <summary>水平スクロールバー</summary>
            HTHSCROLL = 6

            ''' <summary>サイズ変更可能なウィンドウの左端で（ユーザーはマウスをクリックしてウィンドウを水平方向にサイズ変更できます）</summary>
            HTLEFT = 10

            ''' <summary>メニュー</summary>
            HTMENU = 5

            ''' <summary>最大化ボタン</summary>
            HTMAXBUTTON = 9

            ''' <summary>最小化ボタン</summary>
            HTMINBUTTON = 8

            ''' <summary>画面の背景またはウィンドウ間の分割線</summary>
            HTNOWHERE = 0

            ''' <summary>最小化ボタン</summary>
            HTREDUCE = HTMINBUTTON

            ''' <summary>サイズ変更可能なウィンドウの右端で（ユーザーはマウスをクリックしてウィンドウを水平方向にサイズ変更できます）</summary>
            HTRIGHT = 11

            ''' <summary>サイズボックス（HTGROWBOXと同じ）</summary>
            HTSIZE = HTGROWBOX

            ''' <summary>メインウィンドウまたは子ウィンドウ内の[閉じる]ボタンで</summary>
            HTSYSMENU = 3

            ''' <summary>ウィンドウの上部水平境界</summary>
            HTTOP = 12

            ''' <summary>ウィンドウ境界の左上隅にあります</summary>
            HTTOPLEFT = 13

            ''' <summary>ウィンドウ枠の右上隅にあります</summary>
            HTTOPRIGHT = 14

            ''' <summary>
            '''   現在同じスレッド内の別のウィンドウでカバーされているウィンドウ（メッセージは、
            '''   HTTRANSPARENT以外のコードを返すまで同じスレッドの下にあるウィンドウに送信されます）
            ''' </summary>
            HTTRANSPARENT = (-1)

            ''' <summary>垂直スクロールバー</summary>
            HTVSCROLL = 7

            ''' <summary>最大化ボタン</summary>
            HTZOOM = HTMAXBUTTON

        End Enum

    End Class

    ''' <summary>SW_SHOWMINIMIZED</summary>
    ''' <remarks>
    '''   ウィンドウをアクティブにして最小化
    ''' </remarks>
    Public Const SW_SHOWMINIMIZED As Integer = 2

    ''' <summary>SW_SHOWNORMAL</summary>
    ''' <remarks>
    '''   ウィンドウをアクティブにして表示かつ、ウィンドウが最小化または最大化されているときは位置とサイズを元に戻す
    ''' </remarks>
    Public Const SW_SHOWNORMAL As Integer = 1

#End Region

#End Region

#Region "列挙体"

    ''' <summary>
    '''   画像の出力形式
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum OutputExtension

        bmp

        png

        gif

        jpg

    End Enum

#End Region

#Region "構造体"

    '------------------------------------------------★構造体メモ★------------------------------------------------'
    '   StructLayoutについて                                                                                       '
    '     名前空間: System.Runtime.InteropServices                                                                 '
    '     アセンブリ: mscorlib                                                                                     '
    '     継承・インターフェイス: Attribute, _Attribute                                                            '
    '     StructLayout属性は、メモリ上でのフィールド(メンバ変数)の配置方法を指定するための属性です                 '
    '   LayoutKind.Sequentialについて                                                                              '
    '     ランタイムによる自動的な並べ替えを行わず、コード上で記述されている順序のままフィールドを配置する         '
    '   ※マネージ環境で用いられる構造体はアンマネージ環境のものと互換性がありません。                             '
    '     相互運用では構造体の定義にStructLayoutアトリビュートを付加することでアンマネージ互換の構造体を定義します '
    '--------------------------------------------------------------------------------------------------------------'

    ''' <summary>POINT構造体</summary>
    ''' <remarks>
    '''     2 次元空間における、x 座標と y 座標の組を表します
    ''' </remarks>
    <Serializable(), StructLayout(LayoutKind.Sequential)> _
    Public Structure POINT

        ''' <summary>X座標</summary>
        ''' <remarks></remarks>
        Public X As Integer

        ''' <summary>Y座標</summary>
        ''' <remarks></remarks>
        Public Y As Integer

        ''' <summary>コンストラクタ</summary>
        ''' <param name="pX">X座標</param>
        ''' <param name="pY">Y座標</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pX As Integer, ByVal pY As Integer)

            Me.X = X
            Me.Y = Y

        End Sub

    End Structure

    ''' <summary>RECT構造体</summary>
    ''' <remarks>
    '''     四角形の左上隅および右下隅の座標を定義します
    ''' </remarks>
    <Serializable(), StructLayout(LayoutKind.Sequential)> _
    Public Structure RECT

        ''' <summary>Left</summary>
        ''' <remarks>四角形の左上隅のＸ座標を指定します</remarks>
        Public Left As Integer

        ''' <summary>Top</summary>
        ''' <remarks>四角形の左上隅のＹ座標を指定します</remarks>
        Public Top As Integer

        ''' <summary>Right</summary>
        ''' <remarks>四角形の右下隅のＸ座標を指定します</remarks>
        Public Right As Integer

        ''' <summary>Bottom</summary>
        ''' <remarks>四角形の右下隅のＹ座標を指定します</remarks>
        Public Bottom As Integer

        ''' <summary>コンストラクタ</summary>
        ''' <param name="pLeft">Left</param>
        ''' <param name="pTop">Top</param>
        ''' <param name="pRight">Right</param>
        ''' <param name="pBottom">Bottom</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pLeft As Integer, ByVal pTop As Integer, ByVal pRight As Integer, ByVal pBottom As Integer)

            Me.Left = pLeft
            Me.Top = pTop
            Me.Right = pRight
            Me.Bottom = pBottom

        End Sub

    End Structure

    ''' <summary>WINDOWPLACEMENT構造体</summary>
    ''' <remarks>
    '''     画面上のウィンドウの配置についての情報が含まれています
    ''' </remarks>
    <Serializable(), StructLayout(LayoutKind.Sequential)> _
    Public Structure WINDOWPLACEMENT

        ''' <summary>length</summary>
        ''' <remarks>
        '''   構造体サイズをバイト単位で指定します。
        ''' </remarks>
        Public length As Integer

        ''' <summary>flags</summary>
        ''' <remarks>
        '''   最小化されたウィンドウとウィンドウを復元する方法の位置を制御するフラグを指定します。
        '''   このメンバーは、次のフラグの一方または両方を指定できます。
        '''     ・WPF_SETMINPOSITION最小化されたウィンドウの x 位置と y 位置を指定できることを示すします。
        '''       このフラグである必要があります指定された座標で設定されているかどうか、 ptMinPositionメンバーです。
        '''     ・WPF_RESTORETOMAXIMIZED復元されたウィンドウが最大化される、最小化する前に、
        '''       最大化されていたかどうかに関係なくを指定します。 この設定は、ウィンドウを閉じたときにのみ有効です。 
        '''       既定の復元動作は変わりません。 このフラグは有効な場合にのみ、このメンバーは値が指定されて、 showCmdメンバーです。
        ''' </remarks>
        Public flags As Integer

        ''' <summary>showCmd</summary>
        ''' <remarks>
        '''   ウィンドウの現在の表示状態を指定します。 このメンバーは、次の値のいずれかになります。
        '''    ・SW_HIDEウィンドウを非表示にし、別のウィンドウをアクティブ化を渡します。
        '''    ・SW_MINIMIZE指定されたウィンドウを最小化し、システムの一覧の最上位ウィンドウを表示します。
        '''    ・SW_RESTOREにアクティブとウィンドウが表示されます。 
        '''      ウィンドウが最小化または最大化されている場合は、Windows が復元され、元のサイズと位置 (同じSW_SHOWNORMAL)。
        '''    ・SW_SHOWウィンドウをアクティブにし、現在のサイズと位置に表示されます。
        '''    ・SW_SHOWMAXIMIZEDウィンドウをアクティブにし、最大化されたウィンドウとして表示されます。
        '''      このメンバーはウィンドウをアクティブにし、アイコンとして表示します。
        '''    ・SW_SHOWMINNOACTIVEウィンドウをアイコンとして表示します。 
        '''      現在アクティブなウィンドウは、アクティブなままです。
        '''    ・SW_SHOWNA現在の状態で、ウィンドウを表示します。 
        '''      現在アクティブなウィンドウは、アクティブなままです。
        '''    ・SW_SHOWNOACTIVATE最新のサイズと位置で、ウィンドウを表示します。 
        '''      現在アクティブなウィンドウは、アクティブなままです。
        '''    ・SW_SHOWNORMALにアクティブとウィンドウが表示されます。 
        '''      ウィンドウが最小化または最大化されている場合は、Windows が復元され、元のサイズと位置 (同じSW_RESTORE)。
        ''' </remarks>
        Public showCmd As Integer

        ''' <summary>minPosition</summary>
        ''' <remarks>
        '''   ウィンドウが最小化されているときは、ウィンドウの左上隅の位置を指定します。
        ''' </remarks>
        Public minPosition As Point

        ''' <summary></summary>
        ''' <remarks>
        '''   ウィンドウを最大化するときは、ウィンドウの左上隅の位置を指定します。
        ''' </remarks>
        Public maxPosition As Point

        ''' <summary></summary>
        ''' <remarks>
        '''   ウィンドウが標準の (復元) の位置にある場合は、ウィンドウの座標を指定します。
        ''' </remarks>
        Public normalPosition As Rect

    End Structure

#End Region

End Class
