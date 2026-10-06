Attribute VB_Name = "ModSplitByArea"
'==============================================================================
' 地区別にフォルダとブックを分けて保存するマクロ(Python 講座の比較用)
' Mac 版 Excel 向け(Windows 版 Excel でもそのまま動く書き方)
'
' やること(Python ノートブックの 3章 + 5章 と同じ):
'   1. 年度シート(2023年度 / 2024年度 / 2025年度)を、列の並びを揃えて1つに統合
'   2. 地区ごとにフォルダを作り、その中に地区別のブックを保存(20フォルダ・20ファイル)
'   3. かかった秒数をメッセージで表示
'
' 使い方(Mac 版 Excel):
'   1. sales_data.xlsx を Mac 上のふつうのフォルダに置いて Excel で開く
'      ※ OneDrive / iCloud Drive / SharePoint 上のフォルダは避ける
'   2. [ツール] - [マクロ] - [Visual Basic Editor] で VBE を開く
'   3. [挿入] - [標準モジュール] で空のモジュールを作り、このコードを貼り付ける
'   4. Excel の画面に戻って sales_data.xlsx を手前に表示し、
'      [ツール] - [マクロ] - [マクロ...] から SplitByArea を実行する
'   5. 最初に1回だけ「アクセス権を付与」のダイアログが出るので許可する
'      (処理時間は、許可した後から計り始める)
'
' 出力先:
'   sales_data.xlsx と同じフォルダの「地区別_VBA」フォルダ
'     地区別_VBA/東京/東京_売上.xlsx のように保存される
'   ※ そのフォルダに書き込めなかったときは、Excel 専用の作業フォルダに保存し、
'     最後のメッセージに実際の保存先を表示する
'
' ※ sales_data.xlsx 自体は書き換えません
'==============================================================================
Option Explicit

' メインの処理
Sub SplitByArea()

    Dim startTime As Double     ' 開始時刻
    Dim srcBook As Workbook     ' 元データのブック(sales_data.xlsx)
    Dim allBook As Workbook     ' 統合用の一時ブック
    Dim wsAll As Worksheet      ' 統合シート
    Dim wsSrc As Worksheet      ' 読み込み中の年度シート
    Dim wsFirst As Worksheet    ' 列の並びのお手本にするシート(最初の年度シート)
    Dim newBook As Workbook     ' 地区別に保存するブック
    Dim sep As String           ' フォルダの区切り文字(Mac は /、Windows は \)
    Dim outRoot As String       ' 出力先の大もとのフォルダ
    Dim folderPath As String    ' 地区ごとのフォルダ
    Dim areas As Collection     ' 地区名の一覧(重複なし)
    Dim areaName As String      ' 地区名
    Dim headerName As String    ' 列の見出し
    Dim colCount As Long        ' お手本シートの列数
    Dim areaCol As Variant      ' 「地区」列の位置
    Dim yearCol As Long         ' 「年度」列の位置
    Dim srcCol As Variant       ' 元シートでの列の位置
    Dim lastRow As Long         ' 元シートの最終行
    Dim rowCount As Long        ' 元シートのデータ行数
    Dim destRow As Long         ' 統合シートの次に書き込む行
    Dim lastAll As Long         ' 統合シートの最終行
    Dim i As Long
    Dim j As Long
    Dim r As Long

    ' 動作に関わる日本語は、貼り付けやインポートで文字化けしても動くように
    ' 文字コード(ChrW)で組み立てる
    Dim wordYear As String      ' 「年度」
    Dim wordArea As String      ' 「地区」
    Dim wordOutDir As String    ' 「地区別_VBA」
    Dim wordSales As String     ' 「_売上」
    wordYear = ChrW(&H5E74) & ChrW(&H5EA6)
    wordArea = ChrW(&H5730) & ChrW(&H533A)
    wordOutDir = ChrW(&H5730) & ChrW(&H533A) & ChrW(&H5225) & "_VBA"
    wordSales = "_" & ChrW(&H58F2) & ChrW(&H4E0A)

    ' いま手前に表示しているブックを元データとして扱う
    Set srcBook = ActiveWorkbook
    sep = Application.PathSeparator

    ' 年度シート(名前が 2023 のような数字で始まるシート)のうち、最初の1枚をお手本にする
    For Each wsSrc In srcBook.Worksheets
        If Val(wsSrc.Name) > 0 Then
            Set wsFirst = wsSrc
            Exit For
        End If
    Next wsSrc
    If wsFirst Is Nothing Then
        MsgBox "年度シートが見つかりません。sales_data.xlsx を手前に表示してから実行してください。", vbExclamation
        Exit Sub
    End If

    ' 保存先を作れない場所(未保存のブック、OneDrive など)なら中止する
    If srcBook.Path = "" Or LCase(Left(srcBook.Path, 4)) = "http" Then
        MsgBox "sales_data.xlsx を Mac 上のふつうのフォルダに保存してから実行してください。", vbExclamation
        Exit Sub
    End If

    ' Mac 版 Excel だけ: ブックのあるフォルダへのアクセス許可を、最初に1回だけもらう
    ' (これをしないと、ファイルを保存するたびに許可のダイアログが出る)
#If Mac Then
    If Not GrantAccessToMultipleFiles(Array(srcBook.Path)) Then
        MsgBox "フォルダへのアクセスが許可されなかったので中止しました。", vbExclamation
        Exit Sub
    End If
#End If

    ' 出力先の大もとのフォルダを作る
    outRoot = srcBook.Path & sep & wordOutDir
    MakeFolder outRoot

    ' 作れなかったときは、Excel が許可なしで書き込める専用の作業フォルダに切り替える
#If Mac Then
    If Not FolderExists(outRoot) Then
        outRoot = Environ("HOME") & sep & wordOutDir
        MakeFolder outRoot
    End If
#End If
    If Not FolderExists(outRoot) Then
        MsgBox "出力先のフォルダを作れませんでした。" & vbLf & outRoot, vbExclamation
        Exit Sub
    End If

    ' ストップウォッチをスタート(許可のダイアログを閉じた後から計る)
    startTime = Timer

    ' 画面のちらつきと確認メッセージを止める
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '--------------------------------------------------------------------------
    ' 1. 年度シートを、列の並びを揃えて1つのシートに統合する
    '--------------------------------------------------------------------------
    ' 統合用に、シート1枚の新しいブックを作る
    Set allBook = Workbooks.Add(xlWBATWorksheet)
    Set wsAll = allBook.Worksheets(1)

    ' お手本シートの見出し行をそのまま写し、右端に「年度」列を足す
    colCount = wsFirst.Cells(1, wsFirst.Columns.Count).End(xlToLeft).Column
    For j = 1 To colCount
        wsAll.Cells(1, j).Value = wsFirst.Cells(1, j).Value
    Next j
    yearCol = colCount + 1
    wsAll.Cells(1, yearCol).Value = wordYear

    ' 「地区」列が左から何番目かを調べる
    areaCol = Application.Match(wordArea, wsAll.Rows(1), 0)
    If IsError(areaCol) Then
        MsgBox "「地区」列が見つかりません。", vbExclamation
        GoTo Cleanup
    End If

    ' 次に書き込む行(2行目から)
    destRow = 2

    ' 年度シートを1枚ずつ処理する
    For Each wsSrc In srcBook.Worksheets
        If Val(wsSrc.Name) > 0 Then
            lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            rowCount = lastRow - 1

            If rowCount > 0 Then
                ' お手本の列の順に、元シートから同じ見出しの列を探してコピーする
                For j = 1 To colCount
                    headerName = CStr(wsAll.Cells(1, j).Value)
                    srcCol = Application.Match(headerName, wsSrc.Rows(1), 0)
                    If IsError(srcCol) Then
                        MsgBox "シート「" & wsSrc.Name & "」に列「" & headerName & "」がありません。", vbExclamation
                        GoTo Cleanup
                    End If
                    wsSrc.Range(wsSrc.Cells(2, srcCol), wsSrc.Cells(lastRow, srcCol)).Copy _
                        Destination:=wsAll.Cells(destRow, j)
                Next j

                ' 「年度」列に、シート名の先頭の数字(2023 など)を入れる
                wsAll.Range(wsAll.Cells(destRow, yearCol), wsAll.Cells(destRow + rowCount - 1, yearCol)).Value = _
                    CLng(Val(wsSrc.Name))

                destRow = destRow + rowCount
            End If
        End If
    Next wsSrc

    lastAll = destRow - 1

    '--------------------------------------------------------------------------
    ' 2. 地区名の一覧(重複なし)を作る
    '--------------------------------------------------------------------------
    ' Collection は同じキーを2回登録するとエラーになるので、
    ' エラーを無視して登録すれば重複なしの一覧になる
    Set areas = New Collection
    On Error Resume Next
    For r = 2 To lastAll
        areaName = CStr(wsAll.Cells(r, areaCol).Value)
        areas.Add areaName, areaName
    Next r
    On Error GoTo 0

    '--------------------------------------------------------------------------
    ' 3. 地区ごとにフォルダを作り、地区別のブックを保存する
    '--------------------------------------------------------------------------
    For i = 1 To areas.Count
        areaName = areas(i)

        ' 地区のフォルダを作る
        folderPath = outRoot & sep & areaName
        MakeFolder folderPath

        ' 統合シートを、その地区だけに絞り込む(オートフィルター)
        wsAll.Range("A1").CurrentRegion.AutoFilter Field:=areaCol, Criteria1:="=" & areaName

        ' 新しいブックを作り、絞り込んで見えている行だけを貼り付ける
        Set newBook = Workbooks.Add(xlWBATWorksheet)
        wsAll.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy _
            Destination:=newBook.Worksheets(1).Range("A1")
        newBook.Worksheets(1).Columns.AutoFit

        ' 「東京_売上.xlsx」のような名前で保存して閉じる
        newBook.SaveAs Filename:=folderPath & sep & areaName & wordSales & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        newBook.Close SaveChanges:=False
    Next i

    ' 絞り込みを解除する
    wsAll.AutoFilterMode = False

Cleanup:
    ' 統合用の一時ブックを保存せずに閉じる
    allBook.Close SaveChanges:=False

    ' 止めていた設定を元に戻す
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' 途中で中止した場合はここで終わる
    If areas Is Nothing Then Exit Sub

    ' かかった時間と保存先を表示する
    MsgBox areas.Count & " 地区ぶんのフォルダとファイルを作成しました。" & vbLf & _
           "処理時間: " & Format(Timer - startTime, "0.0") & " 秒" & vbLf & _
           "出力先: " & outRoot, vbInformation

End Sub

' フォルダを作る(すでにある場合は何もしない)
' Mac 版の VBA は日本語のフォルダ名を MkDir で作れないため、
' 補助スクリプト PythonKouza.scpt(~/Library/Application Scripts/com.microsoft.Excel/)に作ってもらう
Private Sub MakeFolder(ByVal folderPath As String)
    On Error Resume Next
#If Mac Then
    AppleScriptTask "PythonKouza.scpt", "MakeFolder", folderPath
#Else
    MkDir folderPath
#End If
    On Error GoTo 0
End Sub

' フォルダがあるかを調べる(Mac は日本語のパスを扱えるよう補助スクリプトで調べる)
Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
#If Mac Then
    FolderExists = (AppleScriptTask("PythonKouza.scpt", "FolderExists", folderPath) = "1")
#Else
    FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
#End If
    On Error GoTo 0
End Function
