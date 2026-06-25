Attribute VB_Name = "CConstExtractor"
' =====================================================================
'  CConstExtractor  -  C言語ソースの定数一覧抽出ツール (Excel VBA)
' ---------------------------------------------------------------------
'  C言語ソース(.c / .h)を読み込み、以下の定数を抽出して一覧化する。
'    - #define           (オブジェクト形式マクロ)
'    - const 変数定義     (スカラ / 配列)
'    - enum メンバ
'
'  出力列:  ファイル / 行 / 種別 / 定数名 / 値 / 展開後 / 評価値
'           / 配列要素数 / 参照している定数
'
'  「値に別の定数を使っている」場合は、その定数を再帰的に展開し
'  (展開後 列)、可能なら数値まで評価して(評価値 列)、どの定数を
'  参照しているか(参照している定数 列)も並べて確認できる。
'
'  --- 使い方 (一覧出力) -------------------------------------------
'   1. Excel の VBE を開く (Alt+F11)
'   2. [ファイル] > [ファイルのインポート] で本 .bas を取り込む
'   3. ExtractCConstants を実行 (F5)
'   4. [はい]=フォルダ選択(サブフォルダも再帰) / [いいえ]=ファイル個別選択
'   5. C ファイルごとにシートを分けて出力 (シート名 = ファイル名)
'
'  --- 使い方 (まとめて検索) ---------------------------------------
'   1. 作業中シートの A 列に、調べたい 定数 / 列挙体 の名前を縦に列挙
'   2. その状態で LookupConstants を実行
'   3. 解析データが未読込なら ExtractCConstants と同じ要領でソースを選択
'      (一度 ExtractCConstants 済みなら再選択なしで即検索)
'   4. 各行の B 列以降に
'        種別 / 分類 / 値 / 展開後 / 評価値 / 配列要素数 / 参照 / ファイル / 行
'      が書き込まれる。未定義の名前は B 列に "(該当なし)"。
'
'  --- 制限事項 -----------------------------------------------------
'   * 関数形式マクロ  #define MAX(a,b) ...  は値を評価せず一覧化のみ。
'   * 1行に複数の const を「;」で区切った書き方は最初の文のみ対象。
'     (例: const int a=1; const int b=2;  を1行に書くと b を取りこぼす)
'   * 評価器は + - * / % << >> & | ^ ~ ( ) と
'     10進/16進(0x)/8進(0..) リテラルに対応。sizeof 等は評価不可。
' =====================================================================

' レコードのフィールド位置
Private Const F_FILE As Long = 0
Private Const F_LINE As Long = 1
Private Const F_KIND As Long = 2
Private Const F_NAME As Long = 3
Private Const F_RAW As Long = 4     ' スカラ値の生テキスト
Private Const F_BRK As Long = 5     ' 配列の括弧部 (例 "[3][2]")
Private Const F_INIT As Long = 6    ' 配列の初期化子 (例 "{1,2,3}")

Private mRecords As Collection      ' 各要素は Variant 配列(0..6)
Private mMap As Object              ' 定数名 -> スカラ生テキスト
Private mBase As String             ' 現在処理中ファイルのベース名

' --- 評価器の作業用状態 -------------------------------------------
Private mTok() As String
Private mTT() As String
Private mNT As Long
Private mPos As Long
Private mEvalOK As Boolean


' =====================================================================
'  エントリポイント
' =====================================================================
' C ソースを解析し、ファイル単位のシートに一覧出力する
Public Sub ExtractCConstants()
    If Not LoadSources() Then Exit Sub
    OutputResults
End Sub

' 作業中シートの A 列に並んだ定数名を一括検索し、各行に値を出力する
'   - A 列 : 検索したい 定数 / 列挙体 の名前 (1行1名)
'   - B 列以降 : 解析結果を書き込む
'  解析データが未読み込みなら、その場でソースを読み込む。
Public Sub LookupConstants()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    ' 解析データが無ければ読み込む
    If mRecords Is Nothing Then
        If Not LoadSources() Then Exit Sub
    ElseIf mRecords.Count = 0 Then
        If Not LoadSources() Then Exit Sub
    End If

    ' 定数名 -> レコード の索引を作る (重複名は最初の定義を優先)
    Dim idx As Object: Set idx = NewDict()
    Dim rec As Variant
    For Each rec In mRecords
        Dim nm As String: nm = CStr(rec(F_NAME))
        If Len(nm) > 0 Then
            If Not idx.Exists(nm) Then idx.Add nm, rec
        End If
    Next rec

    ' A 列の最終行
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row   ' xlUp = -4162
    If lastRow < 1 Then Exit Sub

    ' 結果列の見出し (B 列以降)。検索キーは A 列にそのまま残す。
    Dim heads As Variant
    heads = Array("種別", "分類", "値", "展開後", "評価値", "配列要素数", "参照している定数", "ファイル", "行")

    Application.ScreenUpdating = False
    Dim foundN As Long, missN As Long
    Dim r As Long
    For r = 1 To lastRow
        Dim key As String: key = Trim(CStr(ws.Cells(r, 1).Value))
        ' 見出し行らしき "定数名" はスキップ
        If Len(key) = 0 Or key = "定数名" Then
            ' なにもしない
        ElseIf idx.Exists(key) Then
            Dim f As Variant: f = ComputeFields(idx(key))
            ' f = (file,line,kind,category,name,value,resolved,eval,count,refs)
            ws.Cells(r, 2).Value = f(2)    ' 種別
            ws.Cells(r, 3).Value = f(3)    ' 分類
            ws.Cells(r, 4).Value = f(5)    ' 値
            ws.Cells(r, 5).Value = f(6)    ' 展開後
            ws.Cells(r, 6).Value = f(7)    ' 評価値
            ws.Cells(r, 7).Value = f(8)    ' 配列要素数
            ws.Cells(r, 8).Value = f(9)    ' 参照している定数
            ws.Cells(r, 9).Value = f(0)    ' ファイル
            ws.Cells(r, 10).Value = f(1)   ' 行
            foundN = foundN + 1
        Else
            ws.Cells(r, 2).Value = "(該当なし)"
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 10)).ClearContents
            missN = missN + 1
        End If
    Next r

    ' 見出し行を 1 行目の上に挿入 (A1 がデータなら退避)
    PlaceLookupHeader ws, heads

    ws.Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "検索完了: ヒット " & foundN & " 件 / 該当なし " & missN & " 件", _
           vbInformation, "CConstExtractor"
End Sub

' ルックアップ結果の見出しを 1 行目に用意する。
' 既に A1 が見出し("定数名") ならその行へ、そうでなければ 1 行挿入して付ける。
Private Sub PlaceLookupHeader(ws As Worksheet, heads As Variant)
    Dim hasHeader As Boolean
    hasHeader = (Trim(CStr(ws.Cells(1, 1).Value)) = "定数名")
    If Not hasHeader Then
        ws.Rows(1).Insert
        ws.Cells(1, 1).Value = "定数名"
    End If
    Dim c As Long
    For c = 0 To UBound(heads)
        ws.Cells(1, c + 2).Value = heads(c)
    Next c
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(heads) + 2)).Font.Bold = True
End Sub

' 入力方法を選んでソースを読み込み mRecords / mMap を構築する。
' 中断・対象なしなら False。
Private Function LoadSources() As Boolean
    LoadSources = False

    ' 入力方法を選択 (はい=フォルダ / いいえ=ファイル個別選択)
    Dim ans As VbMsgBoxResult
    ans = MsgBox("フォルダ内の C ソースをまとめて解析しますか?" & vbCrLf & _
                 "[はい] = フォルダを選択 (サブフォルダも再帰的に対象)" & vbCrLf & _
                 "[いいえ] = ファイルを個別に選択", _
                 vbYesNoCancel + vbQuestion, "CConstExtractor")
    If ans = vbCancel Then Exit Function

    ' 解析対象パスを集める
    Dim paths As Collection: Set paths = New Collection
    If ans = vbYes Then
        Dim folder As String: folder = PickFolder()
        If Len(folder) = 0 Then Exit Function
        Set paths = CollectFiles(folder)
        If paths.Count = 0 Then
            MsgBox "指定フォルダに .c/.h/.cpp/.hpp などが見つかりませんでした。", vbExclamation, "CConstExtractor"
            Exit Function
        End If
    Else
        Dim files As Variant
        files = Application.GetOpenFilename( _
            "C source (*.c;*.h;*.cpp;*.hpp),*.c;*.h;*.cpp;*.hpp,All files (*.*),*.*", _
            , "解析する C ソースを選択してください", , True)
        If VarType(files) = vbBoolean Then
            If files = False Then Exit Function   ' キャンセル
        End If
        If IsArray(files) Then
            Dim t As Long
            For t = LBound(files) To UBound(files)
                paths.Add CStr(files(t))
            Next t
        Else
            paths.Add CStr(files)
        End If
    End If

    Set mRecords = New Collection
    Set mMap = CreateObject("Scripting.Dictionary")   ' 大文字小文字を区別(C準拠)

    ' 各ファイルをコメント除去して保持
    Dim texts As Collection: Set texts = New Collection
    Dim pv As Variant
    For Each pv In paths
        Dim s As String
        s = StripComments(ReadFile(CStr(pv)))
        texts.Add Array(CStr(pv), s)
    Next pv

    ' パスは依存関係があるため順序が大事:
    '   1) #define を先に全部 map へ
    '   2) enum (メンバ値の計算に define を参照する場合がある)
    '   3) const (配列サイズに define/enum を参照する場合がある)
    Dim it As Variant
    For Each it In texts
        ParseDefines CStr(it(1)), CStr(it(0))
    Next it
    For Each it In texts
        ParseEnums CStr(it(1)), CStr(it(0))
    Next it
    For Each it In texts
        ParseConsts CStr(it(1)), CStr(it(0))
    Next it

    LoadSources = True
End Function


' フォルダ選択ダイアログ (キャンセル時は "")
Private Function PickFolder() As String
    Dim fd As Object
    Set fd = Application.FileDialog(4)   ' 4 = msoFileDialogFolderPicker
    fd.Title = "解析する C ソースのフォルダを選択してください"
    If fd.Show = -1 Then
        PickFolder = fd.SelectedItems(1)
    Else
        PickFolder = ""
    End If
End Function

' フォルダ配下の C ソースを再帰的に収集
Private Function CollectFiles(folder As String) As Collection
    Dim res As Collection: Set res = New Collection
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folder) Then RecurseFolder fso, fso.GetFolder(folder), res
    Set CollectFiles = res
End Function

Private Sub RecurseFolder(fso As Object, fld As Object, res As Collection)
    Dim f As Object
    For Each f In fld.Files
        Dim ext As String: ext = LCase(fso.GetExtensionName(f.Name))
        Select Case ext
            Case "c", "h", "cpp", "hpp", "cc", "hh", "cxx", "hxx"
                res.Add f.Path
        End Select
    Next f
    Dim sub_ As Object
    For Each sub_ In fld.SubFolders
        RecurseFolder fso, sub_, res
    Next sub_
End Sub

' =====================================================================
'  パース: #define
' =====================================================================
Private Sub ParseDefines(s As String, fileName As String)
    Dim re As Object
    Set re = NewRegex("#[ \t]*define[ \t]+([A-Za-z_]\w*)(\([^)]*\))?[ \t]*((?:\\\r?\n|[^\r\n])*)")
    Dim m As Object
    For Each m In re.Execute(s)
        Dim nm As String: nm = m.SubMatches(0)
        Dim isFunc As Boolean: isFunc = (Len(m.SubMatches(1)) > 0)
        Dim val As String: val = CleanCont(m.SubMatches(2))
        Dim ln As Long: ln = LineOf(s, m.FirstIndex)
        If isFunc Then
            PushRec fileName, ln, "MACRO()", nm, m.SubMatches(1) & " " & val, "", ""
        Else
            PushRec fileName, ln, "DEFINE", nm, val, "", ""
            If Len(val) > 0 Then mMap(nm) = val
        End If
    Next m
End Sub


' =====================================================================
'  パース: enum
' =====================================================================
Private Sub ParseEnums(s As String, fileName As String)
    Dim re As Object
    Set re = NewRegex("enum\b[ \t\r\n]*([A-Za-z_]\w*)?[ \t\r\n]*\{([\s\S]*?)\}")
    Dim m As Object
    For Each m In re.Execute(s)
        Dim body As String: body = m.SubMatches(1)
        Dim ln As Long: ln = LineOf(s, m.FirstIndex)
        Dim items As Collection: Set items = SplitTopLevel(body, ",")

        Dim counter As Double: counter = 0
        Dim valid As Boolean: valid = True
        Dim itm As Variant
        For Each itm In items
            Dim t As String: t = Trim(CStr(itm))
            If Len(t) > 0 Then
                Dim nm As String, valOut As String
                Dim eq As Long: eq = TopLevelEq(t)
                If eq > 0 Then
                    nm = FirstIdent(Left(t, eq - 1))
                    Dim res As String: res = ResolveExpr(Trim(Mid(t, eq + 1)), NewDict())
                    Dim sv As String: sv = TryEval(res)
                    If Len(sv) > 0 Then
                        counter = Val(sv): valid = True: valOut = FormatNum(counter)
                    Else
                        valid = False: valOut = res
                    End If
                Else
                    nm = FirstIdent(t)
                    If valid Then valOut = FormatNum(counter) Else valOut = "?"
                End If

                If Len(nm) > 0 Then
                    PushRec fileName, ln, "ENUM", nm, valOut, "", ""
                    mMap(nm) = valOut
                End If
                If valid Then counter = counter + 1
            End If
        Next itm
    Next m
End Sub


' =====================================================================
'  パース: const 変数定義
' =====================================================================
Private Sub ParseConsts(s As String, fileName As String)
    Dim re As Object
    Set re = NewRegex("(?:^|[;{}\)\r\n])[ \t]*((?:static[ \t]+|register[ \t]+|extern[ \t]+)*const[ \t]+[\s\S]*?);")
    Dim m As Object
    For Each m In re.Execute(s)
        Dim stmt As String: stmt = m.SubMatches(0)
        If InStr(stmt, "typedef") = 0 Then
            Dim ln As Long: ln = LineOf(s, m.FirstIndex)
            Dim decls As Collection: Set decls = SplitTopLevel(stmt, ",")
            Dim d As Variant
            For Each d In decls
                Dim nm As String, brk As String, ini As String
                ParseDeclarator CStr(d), nm, brk, ini
                If Len(nm) > 0 Then
                    Dim raw As String
                    If Len(brk) = 0 Then raw = ini Else raw = ""
                    PushRec fileName, ln, "CONST", nm, raw, brk, ini
                    If Len(brk) = 0 And Len(ini) > 0 Then mMap(nm) = ini
                End If
            Next d
        End If
    Next m
End Sub

' 宣言子(例 "const int buf[MAX]" や "x = 5")から 名前 / 括弧部 / 初期化子 を取り出す
Private Sub ParseDeclarator(decl As String, ByRef vName As String, ByRef vBrk As String, ByRef vInit As String)
    vName = "": vBrk = "": vInit = ""

    Dim leftPart As String
    Dim eq As Long: eq = TopLevelEq(decl)
    If eq > 0 Then
        leftPart = Trim(Left(decl, eq - 1))
        vInit = Trim(Mid(decl, eq + 1))
    Else
        leftPart = Trim(decl)
    End If

    ' 左辺に "(" があれば関数/関数ポインタ宣言とみなして除外
    If InStr(leftPart, "(") > 0 Then Exit Sub

    ' 括弧部 [..] を抜き出して name 部から取り除く
    Dim reB As Object: Set reB = NewRegex("\[[^\]]*\]")
    Dim mb As Object
    For Each mb In reB.Execute(leftPart)
        vBrk = vBrk & mb.Value
    Next mb
    Dim namePart As String: namePart = reB.Replace(leftPart, " ")

    ' 名前 = 最後の識別子トークン (型修飾子を読み飛ばした結果)
    Dim reId As Object: Set reId = NewRegex("[A-Za-z_]\w*")
    Dim mm As Object
    For Each mm In reId.Execute(namePart)
        vName = mm.Value
    Next mm
End Sub


' =====================================================================
'  値の展開 (別定数を再帰的に置換)
' =====================================================================
Private Function ResolveExpr(expr As String, visited As Object) As String
    If Len(Trim(expr)) = 0 Then ResolveExpr = expr: Exit Function

    Dim re As Object: Set re = NewRegex("[A-Za-z_]\w*")
    Dim out As String, last As Long
    Dim m As Object
    last = 0
    For Each m In re.Execute(expr)
        Dim id As String: id = m.Value
        out = out & Mid(expr, last + 1, m.FirstIndex - last)
        If mMap.Exists(id) And Not visited.Exists(id) Then
            Dim v2 As Object: Set v2 = CloneDict(visited)
            v2(id) = True
            out = out & "(" & ResolveExpr(CStr(mMap(id)), v2) & ")"
        Else
            out = out & id
        End If
        last = m.FirstIndex + m.Length
    Next m
    out = out & Mid(expr, last + 1)
    ResolveExpr = out
End Function

' 式が参照している既知の定数を「名前=評価値」の形で列挙
Private Function BuildRefs(expr As String) As String
    Dim re As Object: Set re = NewRegex("[A-Za-z_]\w*")
    Dim seen As Object: Set seen = NewDict()
    Dim outc As Collection: Set outc = New Collection
    Dim m As Object
    For Each m In re.Execute(expr)
        Dim id As String: id = m.Value
        If mMap.Exists(id) And Not seen.Exists(id) Then
            seen(id) = True
            Dim ev As String: ev = TryEval(ResolveExpr(CStr(mMap(id)), NewDict()))
            If Len(ev) = 0 Then ev = "(評価不可)"
            outc.Add id & "=" & ev
        End If
    Next m
    BuildRefs = JoinColl(outc, "; ")
End Function


' =====================================================================
'  配列要素数の算出
' =====================================================================
Private Function ComputeArrayCount(brk As String, ini As String) As String
    Dim re As Object: Set re = NewRegex("\[([^\]]*)\]")
    Dim total As Double: total = 1
    Dim known As Boolean: known = True
    Dim dims As Long: dims = 0
    Dim m As Object
    For Each m In re.Execute(brk)
        dims = dims + 1
        Dim inner As String: inner = Trim(m.SubMatches(0))
        If Len(inner) = 0 Then
            If dims = 1 And Len(ini) > 0 Then
                total = total * InitElemCount(ini)
            Else
                known = False
            End If
        Else
            Dim sv As String: sv = TryEval(ResolveExpr(inner, NewDict()))
            If Len(sv) = 0 Then known = False Else total = total * Val(sv)
        End If
    Next m

    If dims = 0 Then
        ComputeArrayCount = ""
    ElseIf known Then
        ComputeArrayCount = FormatNum(total)
    Else
        ComputeArrayCount = "?"
    End If
End Function

' 初期化子のトップレベル要素数
Private Function InitElemCount(ini As String) As Double
    Dim t As String: t = Trim(ini)
    If Len(t) = 0 Then InitElemCount = 0: Exit Function

    If Left(t, 1) = """" Then
        ' 文字列リテラル: 文字数 + 終端 NUL
        InitElemCount = StrLitLen(t) + 1
    ElseIf InStr(t, "{") > 0 Then
        Dim inside As String: inside = InnerBraces(t)
        Dim parts As Collection: Set parts = SplitTopLevel(inside, ",")
        Dim cnt As Long, p As Variant
        For Each p In parts
            If Len(Trim(CStr(p))) > 0 Then cnt = cnt + 1
        Next p
        InitElemCount = cnt
    Else
        InitElemCount = 1
    End If
End Function


' =====================================================================
'  式評価器  (再帰下降)
'  対応: + - * / % << >> & | ^ ~ ( )  と 10進/16進(0x)/8進(0..)
' =====================================================================
Private Function TryEval(expr As String) As String
    If Len(Trim(expr)) = 0 Then TryEval = "": Exit Function
    mEvalOK = True
    Tokenize expr
    mPos = 0
    Dim v As Double: v = pBitOr()
    If mPos < mNT Then mEvalOK = False
    If mEvalOK Then TryEval = FormatNum(v) Else TryEval = ""
End Function

Private Function pBitOr() As Double
    Dim v As Double: v = pBitXor()
    Do While mEvalOK And CurOp("|")
        mPos = mPos + 1
        v = BitOp(v, pBitXor(), "|")
    Loop
    pBitOr = v
End Function
Private Function pBitXor() As Double
    Dim v As Double: v = pBitAnd()
    Do While mEvalOK And CurOp("^")
        mPos = mPos + 1
        v = BitOp(v, pBitAnd(), "^")
    Loop
    pBitXor = v
End Function
Private Function pBitAnd() As Double
    Dim v As Double: v = pShift()
    Do While mEvalOK And CurOp("&")
        mPos = mPos + 1
        v = BitOp(v, pShift(), "&")
    Loop
    pBitAnd = v
End Function
Private Function pShift() As Double
    Dim v As Double: v = pAdd()
    Do While mEvalOK And (CurOp("<<") Or CurOp(">>"))
        Dim op As String: op = mTok(mPos): mPos = mPos + 1
        Dim r As Double: r = pAdd()
        If op = "<<" Then v = v * (2 ^ r) Else v = Int(v / (2 ^ r))
    Loop
    pShift = v
End Function
Private Function pAdd() As Double
    Dim v As Double: v = pMul()
    Do While mEvalOK And (CurOp("+") Or CurOp("-"))
        Dim op As String: op = mTok(mPos): mPos = mPos + 1
        Dim r As Double: r = pMul()
        If op = "+" Then v = v + r Else v = v - r
    Loop
    pAdd = v
End Function
Private Function pMul() As Double
    Dim v As Double: v = pUnary()
    Do While mEvalOK And (CurOp("*") Or CurOp("/") Or CurOp("%"))
        Dim op As String: op = mTok(mPos): mPos = mPos + 1
        Dim r As Double: r = pUnary()
        Select Case op
            Case "*": v = v * r
            Case "/"
                If r = 0 Then
                    mEvalOK = False
                ElseIf v = Int(v) And r = Int(r) Then
                    v = Fix(v / r)          ' C の整数除算 (0方向への切り捨て)
                Else
                    v = v / r
                End If
            Case "%"
                If r = 0 Then mEvalOK = False Else v = v - r * Fix(v / r)
        End Select
    Loop
    pMul = v
End Function
Private Function pUnary() As Double
    If CurOp("-") Then
        mPos = mPos + 1: pUnary = -pUnary()
    ElseIf CurOp("+") Then
        mPos = mPos + 1: pUnary = pUnary()
    ElseIf CurOp("~") Then
        mPos = mPos + 1: pUnary = -(pUnary()) - 1
    Else
        pUnary = pPrimary()
    End If
End Function
Private Function pPrimary() As Double
    If mPos < mNT And mTT(mPos) = "num" Then
        pPrimary = Val(mTok(mPos)): mPos = mPos + 1
    ElseIf mPos < mNT And mTT(mPos) = "lp" Then
        mPos = mPos + 1
        Dim v As Double: v = pBitOr()
        If mPos < mNT And mTT(mPos) = "rp" Then mPos = mPos + 1 Else mEvalOK = False
        pPrimary = v
    Else
        mEvalOK = False
        pPrimary = 0
    End If
End Function

Private Function CurOp(s As String) As Boolean
    CurOp = False
    If mPos < mNT Then
        If mTT(mPos) = "op" And mTok(mPos) = s Then CurOp = True
    End If
End Function

' ビット演算 (| & ^) ... 53bit までの非負整数として処理
Private Function BitOp(a As Double, b As Double, op As String) As Double
    Dim aa As Double, bb As Double
    aa = Int(Abs(a)): bb = Int(Abs(b))
    Dim res As Double, bit As Double
    res = 0: bit = 1
    Dim i As Long
    For i = 0 To 52
        Dim ba As Double, bbit As Double, r As Double
        ba = aa - 2 * Int(aa / 2): aa = Int(aa / 2)
        bbit = bb - 2 * Int(bb / 2): bb = Int(bb / 2)
        Select Case op
            Case "|": If ba = 1 Or bbit = 1 Then r = 1 Else r = 0
            Case "&": If ba = 1 And bbit = 1 Then r = 1 Else r = 0
            Case "^": If ba <> bbit Then r = 1 Else r = 0
        End Select
        res = res + bit * r
        bit = bit * 2
        If aa = 0 And bb = 0 Then Exit For
    Next i
    BitOp = res
End Function

Private Sub Tokenize(expr As String)
    ReDim mTok(0 To Len(expr) + 1)
    ReDim mTT(0 To Len(expr) + 1)
    mNT = 0
    Dim n As Long: n = Len(expr)
    Dim i As Long: i = 1
    Do While i <= n
        Dim c As String: c = Mid(expr, i, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then
            i = i + 1
        ElseIf IsDigitCh(c) Or (c = "." And i < n And IsDigitCh(Mid(expr, i + 1, 1))) Then
            Dim v As Double
            ParseNumber expr, i, v
            AddTok CStr(v), "num"
        ElseIf IsAlphaCh(c) Then
            Dim j As Long: j = i
            Do While j <= n And IsIdentCh(Mid(expr, j, 1))
                j = j + 1
            Loop
            AddTok Mid(expr, i, j - i), "id"
            i = j
        Else
            Dim two As String: two = Mid(expr, i, 2)
            If two = "<<" Or two = ">>" Then
                AddTok two, "op": i = i + 2
            ElseIf InStr("+-*/%&|^~", c) > 0 Then
                AddTok c, "op": i = i + 1
            ElseIf c = "(" Then
                AddTok c, "lp": i = i + 1
            ElseIf c = ")" Then
                AddTok c, "rp": i = i + 1
            Else
                AddTok c, "id": i = i + 1   ' 不明文字 -> 評価失敗させる
            End If
        End If
    Loop
End Sub

Private Sub ParseNumber(expr As String, ByRef i As Long, ByRef outVal As Double)
    Dim n As Long: n = Len(expr)
    Dim c2 As String: c2 = Mid(expr, i, 2)
    If c2 = "0x" Or c2 = "0X" Then
        Dim j As Long: j = i + 2
        Dim v As Double: v = 0
        Do While j <= n And IsHexCh(Mid(expr, j, 1))
            v = v * 16 + HexVal(Mid(expr, j, 1))
            j = j + 1
        Loop
        outVal = v: i = j
    Else
        Dim k As Long: k = i
        Dim numstr As String
        Do While k <= n And (IsDigitCh(Mid(expr, k, 1)) Or Mid(expr, k, 1) = ".")
            numstr = numstr & Mid(expr, k, 1)
            k = k + 1
        Loop
        If Len(numstr) > 1 And Left(numstr, 1) = "0" And InStr(numstr, ".") = 0 And IsOctal(numstr) Then
            outVal = OctVal(numstr)
        Else
            outVal = Val(numstr)
        End If
        i = k
    End If
    ' 整数/浮動小数のサフィックスを読み飛ばす
    Do While i <= n And InStr("uUlLfF", Mid(expr, i, 1)) > 0
        i = i + 1
    Loop
End Sub

Private Sub AddTok(t As String, ty As String)
    mTok(mNT) = t: mTT(mNT) = ty: mNT = mNT + 1
End Sub


' =====================================================================
'  出力
' =====================================================================
Private Sub OutputResults()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    If wb Is Nothing Then Set wb = Workbooks.Add

    ' ファイルを出現順に一意化 (C ファイル単位でシートを分ける)
    Dim fileOrder As Collection: Set fileOrder = New Collection
    Dim seenFile As Object: Set seenFile = NewDict()
    Dim rec As Variant
    For Each rec In mRecords
        Dim fp As String: fp = CStr(rec(F_FILE))
        If Not seenFile.Exists(fp) Then
            seenFile(fp) = True
            fileOrder.Add fp
        End If
    Next rec

    Dim usedNames As Object: Set usedNames = NewDict()
    Dim total As Long: total = 0
    Dim firstWs As Worksheet
    Application.ScreenUpdating = False

    Dim fpv As Variant
    For Each fpv In fileOrder
        Dim fp2 As String: fp2 = CStr(fpv)
        Dim shName As String
        shName = UniqueSheetName(SafeSheetName(BaseName(fp2)), usedNames)

        ' 同名シートが既にあれば作り直す
        Application.DisplayAlerts = False
        On Error Resume Next
        wb.Worksheets(shName).Delete
        On Error GoTo 0
        Application.DisplayAlerts = True

        Dim ws As Worksheet
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = shName
        If firstWs Is Nothing Then Set firstWs = ws

        WriteHeader ws

        Dim row As Long: row = 2
        For Each rec In mRecords
            If CStr(rec(F_FILE)) = fp2 Then
                WriteRecord ws, row, rec
                row = row + 1
                total = total + 1
            End If
        Next rec

        ws.Columns.AutoFit
        ws.Rows(1).AutoFilter
    Next fpv

    Application.ScreenUpdating = True
    If Not firstWs Is Nothing Then firstWs.Activate

    MsgBox total & " 件の定数を " & fileOrder.Count & " 個のシートに抽出しました。", _
           vbInformation, "CConstExtractor"
End Sub

Private Sub WriteHeader(ws As Worksheet)
    Dim hdr As Variant
    hdr = Array("ファイル", "行", "種別", "分類", "定数名", "値", "展開後", "評価値", "配列要素数", "参照している定数")
    Dim col As Long
    For col = 0 To UBound(hdr)
        ws.Cells(1, col + 1).Value = hdr(col)
    Next col
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(hdr) + 1)).Font.Bold = True
End Sub

Private Sub WriteRecord(ws As Worksheet, row As Long, rec As Variant)
    Dim f As Variant: f = ComputeFields(rec)
    Dim c As Long
    For c = 0 To UBound(f)
        ws.Cells(row, c + 1).Value = f(c)
    Next c
End Sub

' 1 レコードから表示用フィールドを計算して配列で返す。
' 戻り値: (0)ファイル (1)行 (2)種別 (3)分類 (4)定数名 (5)値
'         (6)展開後 (7)評価値 (8)配列要素数 (9)参照している定数
Private Function ComputeFields(rec As Variant) As Variant
    Dim raw As String, brk As String, ini As String
    raw = CStr(rec(F_RAW)): brk = CStr(rec(F_BRK)): ini = CStr(rec(F_INIT))

    Dim valueDisp As String
    If Len(brk) > 0 Then
        valueDisp = brk
        If Len(ini) > 0 Then valueDisp = valueDisp & " = " & Shorten(ini, 80)
    Else
        valueDisp = raw
    End If

    Dim resolved As String, evalS As String
    If Len(brk) = 0 And Len(raw) > 0 Then
        resolved = ResolveExpr(raw, NewDict())
        evalS = TryEval(resolved)
        If resolved = raw Then resolved = ""   ' 参照が無ければ空に
    Else
        resolved = "": evalS = ""
    End If

    Dim cnt As String
    If Len(brk) > 0 Then cnt = ComputeArrayCount(brk, ini) Else cnt = ""

    Dim refs As String
    refs = BuildRefs(raw & " " & brk & " " & ini)

    ComputeFields = Array( _
        BaseName(CStr(rec(F_FILE))), rec(F_LINE), CStr(rec(F_KIND)), _
        Category(CStr(rec(F_KIND))), CStr(rec(F_NAME)), valueDisp, _
        resolved, evalS, cnt, refs)
End Function

' シート名に使えない文字を除去し 31 文字以内に収める
Private Function SafeSheetName(nm As String) As String
    Dim s As String: s = nm
    Dim bad As Variant: bad = Array(":", "\", "/", "?", "*", "[", "]")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, CStr(bad(i)), "_")
    Next i
    If Len(s) > 31 Then s = Left(s, 31)
    If Len(Trim(s)) = 0 Then s = "Sheet"
    SafeSheetName = s
End Function

' 既に使った名前と衝突しないよう連番を付ける (31 文字制限を守る)
Private Function UniqueSheetName(base As String, used As Object) As String
    Dim cand As String: cand = base
    Dim i As Long: i = 2
    Do While used.Exists(LCase(cand))
        Dim suf As String: suf = "_" & i
        Dim keep As Long: keep = 31 - Len(suf)
        If keep > Len(base) Then keep = Len(base)
        If keep < 1 Then keep = 1
        cand = Left(base, keep) & suf
        i = i + 1
    Loop
    used(LCase(cand)) = True
    UniqueSheetName = cand
End Function


' =====================================================================
'  汎用ヘルパ
' =====================================================================
Private Sub PushRec(f As String, ln As Long, kind As String, nm As String, _
                    raw As String, brk As String, ini As String)
    Dim a(0 To 6) As Variant
    a(F_FILE) = f: a(F_LINE) = ln: a(F_KIND) = kind: a(F_NAME) = nm
    a(F_RAW) = raw: a(F_BRK) = brk: a(F_INIT) = ini
    mRecords.Add a
End Sub

Private Function NewRegex(pat As String) As Object
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.MultiLine = False
    re.Pattern = pat
    Set NewRegex = re
End Function

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Private Function CloneDict(d As Object) As Object
    Dim n As Object: Set n = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In d.Keys
        n(k) = d(k)
    Next k
    Set CloneDict = n
End Function

Private Function ReadFile(path As String) As String
    Dim f As Integer: f = FreeFile
    Open path For Binary As #f
    Dim sz As Long: sz = LOF(f)
    If sz = 0 Then Close #f: ReadFile = "": Exit Function
    Dim b() As Byte
    ReDim b(0 To sz - 1)
    Get #f, , b
    Close #f
    ReadFile = StrConv(b, vbUnicode)
End Function

' コメント除去 (改行は保持して行番号を維持。文字列/文字リテラルは保護)
Private Function StripComments(s As String) As String
    Dim n As Long: n = Len(s)
    Dim i As Long: i = 1
    Dim sb As String
    Dim state As Integer   ' 0=通常 1=// 2=/* */ 3=string 4=char
    Do While i <= n
        Dim c As String: c = Mid(s, i, 1)
        Dim c2 As String: c2 = Mid(s, i, 2)
        Select Case state
            Case 0
                If c2 = "//" Then
                    state = 1: i = i + 2
                ElseIf c2 = "/*" Then
                    state = 2: i = i + 2
                ElseIf c = """" Then
                    state = 3: sb = sb & c: i = i + 1
                ElseIf c = "'" Then
                    state = 4: sb = sb & c: i = i + 1
                Else
                    sb = sb & c: i = i + 1
                End If
            Case 1
                If c = vbLf Then state = 0: sb = sb & vbLf
                i = i + 1
            Case 2
                If c2 = "*/" Then
                    state = 0: i = i + 2
                Else
                    If c = vbLf Then sb = sb & vbLf
                    i = i + 1
                End If
            Case 3
                If c = "\" Then
                    sb = sb & Mid(s, i, 2): i = i + 2
                ElseIf c = """" Then
                    state = 0: sb = sb & c: i = i + 1
                Else
                    sb = sb & c: i = i + 1
                End If
            Case 4
                If c = "\" Then
                    sb = sb & Mid(s, i, 2): i = i + 2
                ElseIf c = "'" Then
                    state = 0: sb = sb & c: i = i + 1
                Else
                    sb = sb & c: i = i + 1
                End If
        End Select
    Loop
    StripComments = sb
End Function

' #define の行継続(\ + 改行)を空白へ
Private Function CleanCont(s As String) As String
    s = Replace(s, "\" & vbCrLf, " ")
    s = Replace(s, "\" & vbLf, " ")
    s = Replace(s, "\" & vbCr, " ")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbCr, " ")
    CleanCont = Trim(s)
End Function

' 0始まり文字位置 firstIndex までの行番号 (1始まり)
Private Function LineOf(s As String, firstIndex As Long) As Long
    Dim sub_ As String: sub_ = Left(s, firstIndex)
    LineOf = Len(sub_) - Len(Replace(sub_, vbLf, "")) + 1
End Function

' トップレベル(括弧/文字列の外)の delim 位置で分割
Private Function SplitTopLevel(s As String, delim As String) As Collection
    Dim res As Collection: Set res = New Collection
    Dim depth As Long, i As Long, n As Long
    Dim cur As String, c As String
    Dim inS As Boolean, inC As Boolean
    n = Len(s)
    i = 1
    Do While i <= n
        c = Mid(s, i, 1)
        If inS Then
            cur = cur & c
            If c = "\" Then
                cur = cur & Mid(s, i + 1, 1): i = i + 1
            ElseIf c = """" Then
                inS = False
            End If
        ElseIf inC Then
            cur = cur & c
            If c = "\" Then
                cur = cur & Mid(s, i + 1, 1): i = i + 1
            ElseIf c = "'" Then
                inC = False
            End If
        Else
            If c = """" Then
                inS = True: cur = cur & c
            ElseIf c = "'" Then
                inC = True: cur = cur & c
            ElseIf c = "(" Or c = "[" Or c = "{" Then
                depth = depth + 1: cur = cur & c
            ElseIf c = ")" Or c = "]" Or c = "}" Then
                depth = depth - 1: cur = cur & c
            ElseIf depth = 0 And c = delim Then
                res.Add cur: cur = ""
            Else
                cur = cur & c
            End If
        End If
        i = i + 1
    Loop
    res.Add cur
    Set SplitTopLevel = res
End Function

' トップレベルの単独 '=' の位置 (== <= >= != は除外)。無ければ 0
Private Function TopLevelEq(s As String) As Long
    Dim depth As Long, i As Long, n As Long
    Dim inS As Boolean, inC As Boolean
    n = Len(s)
    For i = 1 To n
        Dim c As String: c = Mid(s, i, 1)
        If inS Then
            If c = "\" Then
                i = i + 1
            ElseIf c = """" Then
                inS = False
            End If
        ElseIf inC Then
            If c = "\" Then
                i = i + 1
            ElseIf c = "'" Then
                inC = False
            End If
        Else
            If c = """" Then
                inS = True
            ElseIf c = "'" Then
                inC = True
            ElseIf c = "(" Or c = "[" Or c = "{" Then
                depth = depth + 1
            ElseIf c = ")" Or c = "]" Or c = "}" Then
                depth = depth - 1
            ElseIf depth = 0 And c = "=" Then
                Dim p As String, nx As String
                If i > 1 Then p = Mid(s, i - 1, 1) Else p = ""
                If i < n Then nx = Mid(s, i + 1, 1) Else nx = ""
                If InStr("=<>!", p) = 0 And nx <> "=" Then
                    TopLevelEq = i: Exit Function
                End If
            End If
        End If
    Next i
    TopLevelEq = 0
End Function

Private Function FirstIdent(s As String) As String
    Dim re As Object: Set re = NewRegex("[A-Za-z_]\w*")
    Dim m As Object: Set m = re.Execute(s)
    If m.Count > 0 Then FirstIdent = m(0).Value Else FirstIdent = ""
End Function

Private Function InnerBraces(s As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(s, "{")
    p2 = InStrRev(s, "}")
    If p1 > 0 And p2 > p1 Then InnerBraces = Mid(s, p1 + 1, p2 - p1 - 1) Else InnerBraces = ""
End Function

' 文字列リテラルの文字数 (エスケープ \x は 1文字として数える)
Private Function StrLitLen(t As String) As Long
    Dim p1 As Long, p2 As Long
    p1 = InStr(t, """")
    p2 = InStrRev(t, """")
    If p1 <= 0 Or p2 <= p1 Then StrLitLen = 0: Exit Function
    Dim inner As String: inner = Mid(t, p1 + 1, p2 - p1 - 1)
    Dim re As Object: Set re = NewRegex("\\.")
    inner = re.Replace(inner, "X")
    StrLitLen = Len(inner)
End Function

Private Function JoinColl(c As Collection, sep As String) As String
    Dim out As String, i As Long
    For i = 1 To c.Count
        If i > 1 Then out = out & sep
        out = out & CStr(c(i))
    Next i
    JoinColl = out
End Function

Private Function FormatNum(d As Double) As String
    If d = Int(d) And Abs(d) < 1E+15 Then
        FormatNum = Format(d, "0")
    Else
        FormatNum = CStr(d)
    End If
End Function

Private Function Shorten(s As String, n As Long) As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    If Len(s) > n Then Shorten = Left(s, n) & "..." Else Shorten = s
End Function

' 種別(KIND)から「定数 / 列挙体」の分類を返す
Private Function Category(kind As String) As String
    If kind = "ENUM" Then
        Category = "列挙体"
    Else
        Category = "定数"
    End If
End Function

Private Function BaseName(p As String) As String
    Dim k As Long: k = InStrRev(p, "\")
    If k > 0 Then BaseName = Mid(p, k + 1) Else BaseName = p
End Function

' --- 文字種判定 ----------------------------------------------------
Private Function IsDigitCh(c As String) As Boolean
    IsDigitCh = (c >= "0" And c <= "9")
End Function
Private Function IsAlphaCh(c As String) As Boolean
    IsAlphaCh = (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or c = "_"
End Function
Private Function IsIdentCh(c As String) As Boolean
    IsIdentCh = IsAlphaCh(c) Or IsDigitCh(c)
End Function
Private Function IsHexCh(c As String) As Boolean
    IsHexCh = IsDigitCh(c) Or (c >= "A" And c <= "F") Or (c >= "a" And c <= "f")
End Function
Private Function HexVal(c As String) As Long
    If IsDigitCh(c) Then
        HexVal = Asc(c) - Asc("0")
    ElseIf c >= "a" Then
        HexVal = Asc(c) - Asc("a") + 10
    Else
        HexVal = Asc(c) - Asc("A") + 10
    End If
End Function
Private Function IsOctal(s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        Dim c As String: c = Mid(s, i, 1)
        If c < "0" Or c > "7" Then IsOctal = False: Exit Function
    Next i
    IsOctal = True
End Function
Private Function OctVal(s As String) As Double
    Dim v As Double, i As Long
    For i = 1 To Len(s)
        v = v * 8 + (Asc(Mid(s, i, 1)) - Asc("0"))
    Next i
    OctVal = v
End Function
