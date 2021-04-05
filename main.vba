Option Compare Database
Option Explicit

'メール本文から会社名を推測する
Private Function getCompanyName(var As String) As String

    If InStr(var, "CodeBrew") <> 0 Then
        getCompanyName = "合同会社CodeBrew"
    ElseIf InStr(var, "Classmethod") <> 0 Then
        getCompanyName = "Classmethod株式会社"
    Else
        getCompanyName = ""
    End If

End Function

'メール本文から商品名を推測する
Private Function getProducts() As String()
    Dim db As Database
    Dim qds As QueryDefs
    Dim rs As Recordset
    Dim products() As String
    Dim i As Integer

    Set db = CurrentDb()
    i = 0

    ReDim products(100) '100は商品数

    Set rs = db.OpenRecordset("Q_商品一覧", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!仕切り価格 = True Then
            products(i) = rs!型式
            i = i + 1
        End If
        rs.MoveNext
    Loop
    rs.Close

    getProducts = products
End Function

'メール本文から商品名を推測する
Private Function getAssumedProducts() As String()
    Dim products() As String
    Dim product As Variant
    Dim items() As String
    Dim i As Integer

    products = getProducts()

    ReDim items(5)

    i = 0
    For Each product In products
        If InStr(Me![目次], product) <> 0 And product <> "" Then
            items(i) = product
            i = i + 1
        End If
    Next

    getAssumedProducts = items
End Function

'商品名を抽出できたら1を、それ以外は0を返す
Private Function getDefaultQuantity(hasProductName As Boolean) As String
    If hasProductName Then
        getDefaultQuantity = 1
    Else
        getDefaultQuantity = 0
    End If
End Function

'Officeのバージョンを返す
Private Function getOfficeVersion() As String
    getOfficeVersion = Application.Version
End Function

Private Sub Form_Current()

    Dim companyName As String
    Dim items() As String
    Dim item As Variant
    Dim i As Integer
    Dim db As Database
    Dim rs As Recordset

    companyName = getCompanyName(Me![目次])
    items = getAssumedProducts()

    Set db = CurrentDb()

    '見積を求められている商品の在庫数を取得
    i = 0
    For Each item In items
        Set rs = db.OpenRecordset("Q_商品一覧", dbOpenDynaset)
        Do While Not rs.EOF
            If rs!型式 = item And rs!仕切り価格 = True Then
                items(i) = item & " (在庫数: " & rs!在庫数 & ")"
            End If
            rs.MoveNext
        Loop
        i = i + 1
    Next

    Me![会社名] = companyName
    Me![商品1] = items(0)
    Me![商品2] = items(1)
    Me![商品3] = items(2)
    Me![商品4] = items(3)
    Me![商品5] = items(4)
    Me![商品6] = Null
    Me![数量1] = getDefaultQuantity(Len(items(0)) > 0)
    Me![数量2] = getDefaultQuantity(Len(items(1)) > 0)
    Me![数量3] = getDefaultQuantity(Len(items(2)) > 0)
    Me![数量4] = getDefaultQuantity(Len(items(3)) > 0)
    Me![数量5] = getDefaultQuantity(Len(items(4)) > 0)
    Me![数量6] = 0
    Me![取扱なし].Value = False

End Sub

'型式から商品名を取得する
Private Function getProductNameFromModelName(modelName As String) As String
    Dim db As Database
    Dim qds As QueryDefs
    Dim rs As Recordset
    Dim productName As String

    Set db = CurrentDb()

    Set rs = db.OpenRecordset("Q_商品一覧", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!型式 = modelName And rs!仕切り価格 = False Then
            productName = rs!製品名
        End If
        rs.MoveNext
    Loop
    rs.Close

    getProductNameFromModelName = productName
End Function

'複数台割引をする
Private Function getPriceByRecordsetAndQuantity(rs As Recordset, quantity As Long) As String
    If rs!複数台価格 = True Then
        If quantity < 5 Then
            getPriceByRecordsetAndQuantity = rs!標準価格
        ElseIf quantity < 10 Then
            getPriceByRecordsetAndQuantity = rs![価格（5式以上）]
        ElseIf quantity < 30 Then
            getPriceByRecordsetAndQuantity = rs![価格（10式以上）]
        ElseIf quantity < 50 Then
            getPriceByRecordsetAndQuantity = rs![価格（30式以上）]
        ElseIf quantity < 100 Then
            getPriceByRecordsetAndQuantity = rs![価格（50式以上）]
        End If
    Else
        getPriceByRecordsetAndQuantity = rs!標準価格
    End If
End Function

Private Sub Form_Load()
    DoCmd.GoToRecord acDataForm, "F_見積もり", acLast
End Sub

Private Sub 見積りメール作成_Click()

    If Me![取扱なし].Value = False Then

        If Me![数量1] = 0 And Me![数量2] = 0 And Me![数量3] = 0 And Me![数量4] = 0 And Me![数量5] = 0 And Me![数量6] = 0 Then
            MsgBox ("商品の数量が0です。")
            Exit Sub
        End If

        If Me![数量1] >= 100 Or Me![数量2] >= 100 Or Me![数量3] >= 100 Or Me![数量4] >= 100 Or Me![数量5] >= 100 Or Me![数量6] >= 100 Then
            MsgBox ("商品の数量が100を越える見積メール生成は実装されていません。")
            Exit Sub
        End If

        If Me![数量1] <> 0 And Me![商品1] = "" Then
            MsgBox ("見積もりできない商品が含まれています：商品1。")
            Exit Sub
        ElseIf Me![数量2] <> 0 And Me![商品2] = "" Then
            MsgBox ("見積もりできない商品が含まれています：商品2。")
            Exit Sub
        ElseIf Me![数量3] <> 0 And Me![商品3] = "" Then
            MsgBox ("見積もりできない商品が含まれています：商品3。")
            Exit Sub
        ElseIf Me![数量4] <> 0 And Me![商品4] = "" Then
            MsgBox ("見積もりできない商品が含まれています：商品4。")
            Exit Sub
        ElseIf Me![数量5] <> 0 And Me![商品5] = "" Then
            MsgBox ("見積もりできない商品が含まれています：商品5。")
            Exit Sub
        End If

        If Me![数量6] <> 0 And IsNull(Me![商品6]) Then
            MsgBox ("見積もりできない商品が含まれています：商品6。")
            Exit Sub
        End If

    End If


    Dim companyName As String
    Dim products() As String
    Dim rep As String
    Dim receivedDate As Date
    Dim greeting As String
    Dim ending As String
    Dim estimateIntro As String
    Dim estimateMain As String
    Dim estimateLast As String
    Dim originalPrice As String
    Dim wholosalePrice As String
    Dim conjunctionPrice As Long
    Dim product As String
    Dim quantity As Long

    Dim items() As String
    Dim item As Variant
    Dim db As Database
    Dim rs As Recordset

    companyName = getCompanyName(Me![目次])
    products = getProducts()
    rep = Me![差出人]
    receivedDate = Me![受信日時]
    items = getAssumedProducts()

    Dim name As Variant

    name = Split(rep, " ")
    If name(0) = rep Then
        name = Split(rep, "　")
    End If

    rep = name(0)


    greeting = "いつもお世話になっております。" & vbCrLf & "以下、御見積です｡ "
    ending = "納期は、受注後1週間です。"
    estimateIntro = "*** 御見積 ***" & vbCrLf & vbTab & "　標準価格" & vbTab & "仕切り価格"
    estimateLast = "税別価格" & vbCrLf & "********* "
    estimateMain = ""

    Set db = CurrentDb()

    conjunctionPrice = 0

    If Me![商品1] <> "" And Me![数量1] <> 0 Then
        product = items(0)
        quantity = Me![数量1]

        Set rs = db.OpenRecordset("Q_商品一覧", dbOpenDynaset)
        Do While Not rs.EOF
            If rs!型式 = product And rs!仕切り価格 = False Then
                originalPrice = getPriceByRecordsetAndQuantity(rs, quantity)
            ElseIf rs!型式 = product And rs!仕切り価格 = True Then
                wholosalePrice = getPriceByRecordsetAndQuantity(rs, quantity)
            End If
            rs.MoveNext
        Loop

        estimateMain = product & vbTab & "@" & Format(originalPrice, "#,##0") & "円" & vbTab & "@" & Format(wholosalePrice, "#,##0") & "円"
        conjunctionPrice = Val(wholosalePrice) * quantity
    End If

    If Me![商品2] <> "" And Me![数量2] <> 0 Then
        product = items(1)
        quantity = Me![数量2]

        Set rs = db.OpenRecordset("Q_商品一覧", dbOpenDynaset)
        Do While Not rs.EOF
            If rs!型式 = product And rs!仕切り価格 = False Then
                originalPrice = getPriceByRecordsetAndQuantity(rs, quantity)
            ElseIf rs!型式 = product And rs!仕切り価格 = True Then
                wholosalePrice = getPriceByRecordsetAndQuantity(rs, quantity)
            End If
            rs.MoveNext
        Loop

        If Len(estimateMain) > 0 Then
            estimateMain = estimateMain & vbCrLf
        End If
        estimateMain = estimateMain & product & vbTab & "@" & Format(originalPrice, "#,##0") & "円" & vbTab & "@" & Format(wholosalePrice, "#,##0") & "円"
        conjunctionPrice = conjunctionPrice + Val(wholosalePrice) * quantity
    End If

    If Me![商品3] <> "" And Me![数量3] <> 0 Then
        product = items(2)
        quantity = Me![数量3]

        Set rs = db.OpenRecordset("Q_商品一覧", dbOpenDynaset)
        Do While Not rs.EOF
            If rs!型式 = product And rs!仕切り価格 = False Then
                originalPrice = getPriceByRecordsetAndQuantity(rs, quantity)
            ElseIf rs!型式 = product And rs!仕切り価格 = True Then
                wholosalePrice = getPriceByRecordsetAndQuantity(rs, quantity)
            End If
            rs.MoveNext
        Loop

        If Len(estimateMain) > 0 Then
            estimateMain = estimateMain & vbCrLf
        End If
        estimateMain = estimateMain & product & vbTab & "@" & Format(originalPrice, "#,##0") & "円" & vbTab & "@" & Format(wholosalePrice, "#,##0") & "円"
        conjunctionPrice = conjunctionPrice + Val(wholosalePrice) * quantity
    End If

    If Me![商品4] <> "" And Me![数量4] <> 0 Then
        product = items(3)
        quantity = Me![数量4]

        Set rs = db.OpenRecordset("Q_商品一覧", dbOpenDynaset)
        Do While Not rs.EOF
            If rs!型式 = product And rs!仕切り価格 = False Then
                originalPrice = getPriceByRecordsetAndQuantity(rs, quantity)
            ElseIf rs!型式 = product And rs!仕切り価格 = True Then
                wholosalePrice = getPriceByRecordsetAndQuantity(rs, quantity)
            End If
            rs.MoveNext
        Loop

        If Len(estimateMain) > 0 Then
            estimateMain = estimateMain & vbCrLf
        End If
        estimateMain = estimateMain & product & vbTab & "@" & Format(originalPrice, "#,##0") & "円" & vbTab & "@" & Format(wholosalePrice, "#,##0") & "円"
        conjunctionPrice = conjunctionPrice + Val(wholosalePrice) * quantity
    End If

    If Me![商品5] <> "" And Me![数量5] <> 0 Then
        product = items(4)
        quantity = Me![数量5]

        Set rs = db.OpenRecordset("Q_商品一覧", dbOpenDynaset)
        Do While Not rs.EOF
            If rs!型式 = product And rs!仕切り価格 = False Then
                originalPrice = getPriceByRecordsetAndQuantity(rs, quantity)
            ElseIf rs!型式 = product And rs!仕切り価格 = True Then
                wholosalePrice = getPriceByRecordsetAndQuantity(rs, quantity)
            End If
            rs.MoveNext
        Loop

        If Len(estimateMain) > 0 Then
            estimateMain = estimateMain & vbCrLf
        End If
        estimateMain = estimateMain & product & vbTab & "@" & Format(originalPrice, "#,##0") & "円" & vbTab & "@" & Format(wholosalePrice, "#,##0") & "円"
        conjunctionPrice = conjunctionPrice + Val(wholosalePrice) * quantity
    End If

    If Not IsNull(Me![商品6]) And Me![数量6] <> 0 Then
        product = Me![商品6]
        quantity = Me![数量6]

        Set rs = db.OpenRecordset("Q_商品一覧", dbOpenDynaset)
        Do While Not rs.EOF
            If rs!型式 = product And rs!仕切り価格 = False Then
                originalPrice = getPriceByRecordsetAndQuantity(rs, quantity)
            ElseIf rs!型式 = product And rs!仕切り価格 = True Then
                wholosalePrice = getPriceByRecordsetAndQuantity(rs, quantity)
            End If
            rs.MoveNext
        Loop

        If Len(estimateMain) > 0 Then
            estimateMain = estimateMain & vbCrLf
        End If
        estimateMain = estimateMain & getProductNameFromModelName(product) & vbTab & "@" & Format(originalPrice, "#,##0") & "円" & vbTab & "@" & Format(wholosalePrice, "#,##0") & "円"
        conjunctionPrice = conjunctionPrice + Val(wholosalePrice) * quantity
    End If

    estimateMain = estimateMain & vbCrLf & "合価：" & Format(conjunctionPrice, "#,##0") & "円"

    Dim accountName As String
    Dim objOutlook, ns, inbox, subFolder
    Dim mailman As String
    Dim folder
    Dim mails As Outlook.items
    Dim objReply As Outlook.MailItem
    Dim senderEmailAddress As String
    Dim i As Integer

    accountName = "your-address@domain.co.jp"
    Set objOutlook = CreateObject("Outlook.Application")
    Set ns = objOutlook.getnamespace("MAPI")

    '社内で複数のOfficeを使っている場合、必要に応じて動作を分ける
    If getOfficeVersion() = "16.0" Then 'Office2019
        Set folder = ns.Folders(accountName).Folders.item("見積もり")
    ElseIf getOfficeVersion() = "12.0" Then 'Office2007
        Set folder = ns.Folders(1).Folders(accountName).Folders.item("見積もり")
    Else
        MsgBox "このオフィスのバージョンはサポートされていません。version: " & getOfficeVersion()
        Exit Sub
    End If

    Set mails = folder.items
    mails.Sort "[受信日時]", True

    For i = 1 To mails.Count
        If receivedDate = mails(i).ReceivedTime Then
            senderEmailAddress = mails(i).senderEmailAddress
            Set objReply = mails(i).ReplyAll
    End If
    Next i

    objReply.SendUsingAccount = objOutlook.Session.Accounts.item(accountName)
    objReply.BCC = accountName

    If Me![取扱なし].Value = False Then
        objReply.Body = companyName & vbCrLf & "　" & rep & " 様" & vbCrLf & vbCrLf & greeting & vbCrLf & estimateIntro & vbCrLf & estimateMain & vbCrLf & estimateLast & vbCrLf & ending & objReply.Body
    Else
        objReply.Body = companyName & vbCrLf & "　" & rep & " 様" & vbCrLf & vbCrLf & "いつもお世話になっております。" & vbCrLf & "弊社ではその商品は取り扱いがございません。" & objReply.Body
    End If


    objReply.Display

    Set folder = Nothing

End Sub

Private Function increment(num As Integer) As Integer
    If num + 1 <= 999 Then
        increment = num + 1
    End If
End Function
Private Function decrement(num As Integer) As Integer
    If num - 1 >= 0 Then
        decrement = num - 1
    End If
End Function

Private Sub 数量1を削減_Click()
    Me![数量1] = decrement(Me![数量1])
End Sub

Private Sub 数量1を追加_Click()
    Me![数量1] = increment(Me![数量1])
End Sub

Private Sub 数量2を削減_Click()
    Me![数量2] = decrement(Me![数量2])
End Sub

Private Sub 数量2を追加_Click()
    Me![数量2] = increment(Me![数量2])
End Sub

Private Sub 数量3を削減_Click()
    Me![数量3] = decrement(Me![数量3])
End Sub

Private Sub 数量3を追加_Click()
    Me![数量3] = increment(Me![数量3])
End Sub

Private Sub 数量4を削減_Click()
    Me![数量4] = decrement(Me![数量4])
End Sub

Private Sub 数量4を追加_Click()
    Me![数量4] = increment(Me![数量4])
End Sub

Private Sub 数量5を削減_Click()
    Me![数量5] = decrement(Me![数量5])
End Sub

Private Sub 数量5を追加_Click()
    Me![数量5] = increment(Me![数量5])
End Sub

Private Sub 数量6を削減_Click()
    Me![数量6] = decrement(Me![数量6])
End Sub

Private Sub 数量6を追加_Click()
    Me![数量6] = increment(Me![数量6])
End Sub
