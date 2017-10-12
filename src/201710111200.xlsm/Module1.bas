Attribute VB_Name = "Module1"
Option Explicit


Sub overWorkColorRank()

'csvファイルを読み込む
Dim buf As String 'bufって変数名はきもいので後で置換する
'ファイル場所はあとでボタンから読み込む(とりあえずベタ書き)
Open "C:\Users\t.kawano\Desktop\残業代作成ツール\daily_2017-09-01_2017-10-01.csv" For Input As #1

Do Until EOF(1)
    Line Input #1, buf
    '読み込んだデータをセルに代入する
    Loop
    Close #1

'残業時間カラム(AC)で、時間が存在する部分の背景色を塗る
'残業時間規模によって背景色をかえる
'カラムがなくなったらおわり
'吐き出すお

End Sub


'エクセルにカラム名ぶち込むやつ(いらなくねこれ)
Sub insertExcColumn()
    Dim tmp As Variant
    '↓カラム名がベタだけどたぶん直したい
    tmp = Split("部署コード,部署名,勤務体系,社員コード,社員名,月度,日付,曜日,種別,休暇申請,振替申請,勤務時間変更申請,休日出勤申請,残業申請,早朝勤務申請,遅刻申請,早退申請,日次確定,月次確定,シフト,始業,終業,出社,退社,出社(丸めなし),退社(丸めなし),総労働時間,実労働時間,残業時間,法定休日労働時間,深夜労働時間,欠勤時間,休憩時間,実時間,工数合計時間,備考" _
, ",")
    Range("A1:AJ1").Value = tmp
End Sub
