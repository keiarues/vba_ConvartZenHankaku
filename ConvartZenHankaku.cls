VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConvartZenHankaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'英数
Const alphaNumZenkakuList As String = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ"
Const alphaNumHankakuList As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"

'ASCII記号(「"」ダブルクォーテーションは除く)
Const asciiZenkakuList  As String = "！＃＄％＆’（）＊＋，－．／：；＜＝＞？＠［￥］＾＿‘｛｜｝￣"
Const asciiHankakuList As String = "!#$%&'()*+,-./:;<=>?@[\]^_`{|}~"

'カタカナ清音
Const katakanaSeionZenkakuList As String = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲンァィゥェォャュョッ"
Const katakanaSeionHankakuList  As String = "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝｧｨｩｪｫｬｭｮｯ"

'カタカナ濁音
Const katakanaDakuonZenkakuList As String = "ガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポ"
Const katakanaDakuonHankakuList  As String = "ｶﾞｷﾞｸﾞｹﾞｺﾞｻﾞｼﾞｽﾞｾﾞｿﾞﾀﾞﾁﾞﾂﾞﾃﾞﾄﾞﾊﾞﾋﾞﾌﾞﾍﾞﾎﾞﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ"


'英数とASCII記号
'半角→全角 変換処理
Public Function alphaNumAsciiHankakuToZenkaku(ByVal argStr As String) As String

    Dim reStr As String
    reStr = argStr
    
    reStr = alphaNumHankakuToZenkaku(reStr)
    reStr = asciiHankakuToZenkaku(reStr)
    
    alphaNumAsciiHankakuToZenkaku = reStr
    
End Function


'英数とASCII記号
'全角→半角 変換処理
Public Function alphaNumAsciiZenkakuToHankaku(ByVal argStr As String) As String

    Dim reStr As String
    reStr = argStr
    
    reStr = alphaNumZenkakuToHankaku(reStr)
    reStr = asciiZenkakuToHankaku(reStr)
    
    alphaNumAsciiZenkakuToHankaku = reStr
    
End Function



'カタカナ
'半角→全角 変換処理
Public Function katakanaHankakuToZenkaku(ByVal argStr As String) As String
    
    '戻り値
    Dim reStr As String
    reStr = argStr
    
    'ループ用
    Dim loop1 As Long
    Dim tmpZenSrt As String
    Dim tmpHanSrt As String
    
    
    '濁音 半濁音の変換
    '半角だと2文字なので別ループで変換する
    '半角→全角のときは必ず清音文字の変換前にやること
    For loop1 = 1 To Len(katakanaDakuonZenkakuList)
        
        '置換リストから文字を1文字切り出し
        tmpZenSrt = Mid(katakanaDakuonZenkakuList, loop1, 1)
        
        '濁音の半角カタカナは2文字で構成されてるのでも2文字単位で切り出し
        tmpHanSrt = Mid(katakanaDakuonHankakuList, (loop1 * 2) - 1, 2)
        
        '置換
        reStr = Replace(reStr, tmpHanSrt, tmpZenSrt)
        
    Next
    
    
    For loop1 = 1 To Len(katakanaSeionZenkakuList)
        
        '置換リストから文字を1文字切り出し
        tmpZenSrt = Mid(katakanaSeionZenkakuList, loop1, 1)
        tmpHanSrt = Mid(katakanaSeionHankakuList, loop1, 1)
        
        '置換
        reStr = Replace(reStr, tmpHanSrt, tmpZenSrt)
        
    Next
    
    
    katakanaHankakuToZenkaku = reStr
    
End Function


'カタカナ
'全角→半角 変換処理
Public Function katakanaZenkakuToHankaku(ByVal argStr As String) As String
    
    '戻り値
    Dim reStr As String
    reStr = argStr
    
    'ループ用
    Dim loop1 As Long
    Dim tmpZenSrt As String
    Dim tmpHanSrt As String
    
    
    For loop1 = 1 To Len(katakanaSeionZenkakuList)
        
        '置換リストから文字を1文字切り出し
        tmpZenSrt = Mid(katakanaSeionZenkakuList, loop1, 1)
        tmpHanSrt = Mid(katakanaSeionHankakuList, loop1, 1)
        
        '置換
        reStr = Replace(reStr, tmpZenSrt, tmpHanSrt)
        
    Next
    
    
    '濁音 半濁音の変換
    '半角にすると2文字になるので別ループで変換する
    For loop1 = 1 To Len(katakanaDakuonZenkakuList)
        
        '置換リストから文字を1文字切り出し
        tmpZenSrt = Mid(katakanaDakuonZenkakuList, loop1, 1)
        
        '濁音の半角カタカナは2文字で構成されてるのでも2文字単位で切り出し
        tmpHanSrt = Mid(katakanaDakuonHankakuList, (loop1 * 2) - 1, 2)
        
        '置換
        reStr = Replace(reStr, tmpZenSrt, tmpHanSrt)
        
    Next
    
    
    katakanaZenkakuToHankaku = reStr
    
End Function


'----------------------

'英数
'半角→全角 変換処理
Private Function alphaNumHankakuToZenkaku(ByVal argStr As String) As String
    
    '戻り値
    Dim reStr As String
    reStr = argStr
    
    'ループ用
    Dim loop1 As Long
    Dim tmpZenSrt As String
    Dim tmpHanSrt As String
    
    
    For loop1 = 1 To Len(alphaNumZenkakuList)
        
        '置換リストから文字を1文字切り出し
        tmpZenSrt = Mid(alphaNumZenkakuList, loop1, 1)
        tmpHanSrt = Mid(alphaNumHankakuList, loop1, 1)
        
        '置換
        reStr = Replace(reStr, tmpHanSrt, tmpZenSrt)
        
    Next
    
    
    alphaNumHankakuToZenkaku = reStr
    
End Function


'英数
'全角→半角 変換処理
Private Function alphaNumZenkakuToHankaku(ByVal argStr As String) As String
    
    '戻り値
    Dim reStr As String
    reStr = argStr
    
    'ループ用
    Dim loop1 As Long
    Dim tmpZenSrt As String
    Dim tmpHanSrt As String
    
    
    For loop1 = 1 To Len(alphaNumZenkakuList)
        
        '置換リストから文字を1文字切り出し
        tmpZenSrt = Mid(alphaNumZenkakuList, loop1, 1)
        tmpHanSrt = Mid(alphaNumHankakuList, loop1, 1)
        
        '置換
        reStr = Replace(reStr, tmpZenSrt, tmpHanSrt)
        
    Next

    alphaNumZenkakuToHankaku = reStr
    
End Function



'ASCII記号
'半角→全角 変換処理
Private Function asciiHankakuToZenkaku(ByVal argStr As String) As String
    
    '戻り値
    Dim reStr As String
    reStr = argStr
    
    'ループ用
    Dim loop1 As Long
    Dim tmpZenSrt As String
    Dim tmpHanSrt As String
    
    
    For loop1 = 1 To Len(asciiZenkakuList)
        
        '置換リストから文字を1文字切り出し
        tmpZenSrt = Mid(asciiZenkakuList, loop1, 1)
        tmpHanSrt = Mid(asciiHankakuList, loop1, 1)
        
        '置換
        reStr = Replace(reStr, tmpHanSrt, tmpZenSrt)
        
    Next
    
    
    'ダブルクォーテーションの変換
    reStr = doubleQuotationHankakuToZenkaku(reStr)
    
    
    asciiHankakuToZenkaku = reStr
    
End Function


'ASCII記号
'全角→半角 変換処理
Private Function asciiZenkakuToHankaku(ByVal argStr As String) As String
    
    '戻り値
    Dim reStr As String
    reStr = argStr
    
    'ループ用
    Dim loop1 As Long
    Dim tmpZenSrt As String
    Dim tmpHanSrt As String
    
    
    For loop1 = 1 To Len(asciiZenkakuList)
        
        '置換リストから文字を1文字切り出し
        tmpZenSrt = Mid(asciiZenkakuList, loop1, 1)
        tmpHanSrt = Mid(asciiHankakuList, loop1, 1)
        
        '置換
        reStr = Replace(reStr, tmpZenSrt, tmpHanSrt)
        
    Next


    'ダブルクォーテーションの変換
    reStr = doubleQuotationZenkakuToHankaku(reStr)
    
    
    asciiZenkakuToHankaku = reStr
    
End Function




'ダブルクォーテーション
'半角→全角 変換処理
'全角のダブルクォーテーションがVBAでまともに取り扱えないので専用処理にする
Private Function doubleQuotationHankakuToZenkaku(ByVal argStr As String) As String
    
    '戻り値
    Dim reStr As String
    reStr = argStr
    
    Dim dqZenkaku As String
    Dim dqHankaku As String
    
    
    '全角ダブルクォーテーションの文字コード
    '文字コードだと取り扱い可能になる
    dqZenkaku = Chr(&H8168)
    dqHankaku = Chr(34)
    
    reStr = Replace(reStr, dqHankaku, dqZenkaku)
    
    doubleQuotationHankakuToZenkaku = reStr
    
End Function



'ダブルクォーテーション
'全角→半角 変換処理
'全角のダブルクォーテーションがVBAでまともに取り扱えないので専用処理にする
Private Function doubleQuotationZenkakuToHankaku(ByVal argStr As String) As String
    
    '戻り値
    Dim reStr As String
    reStr = argStr
    
    Dim dqZenkaku As String
    Dim dqHankaku As String
    
    
    '全角ダブルクォーテーションの文字コード
    '文字コードだと取り扱い可能になる
    dqZenkaku = Chr(&H8168)
    dqHankaku = Chr(34)
    
    reStr = Replace(reStr, dqZenkaku, dqHankaku)
    
    doubleQuotationZenkakuToHankaku = reStr
    
End Function








