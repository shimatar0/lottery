Option Explicit Randomize

Dim user_num Dim win_num Dim password

password = InputBox("") user_num = InputBox("くじを引く人数を入力してください") win_num = InputBox("当たりの数を入力してください")

Dim intCounter_1 Dim intCounter_2

ReDim intwinNumbers(user_num)

For intCounter_1 = 0 To user_num Step 1 intwinNumbers(intCounter_1) = 0 Next

Dim intwinnum_count_1 Dim r

intwinnum_count_1 = 0

Dim flag Dim win_num_3

win_num_3 = CInt(win_num)

flag = 0

do while flag <> win_num_3 r = Int(Rnd * user_num) If Not intwinNumbers(r) = 1 Then intwinNumbers(r) = 1 flag = flag + 1 End If loop

intwinnum_count_1 = intCounter_1 + 1

Dim intCounter_3

Dim rand_num Dim find_num() Dim intCounter_4 Dim win_count,win_num_2 Dim intCounter_5 Dim intLoop_cnt

ReDim find_num(user_num)

For intCounter_4 = 0 To user_num Step 1 find_num(intCounter_4) = 0 Next

win_count = 0 intLoop_cnt = 1 win_num_2 = CInt (win_num)

Dim msg Dim debag_win

debag_win = 4

msg = "くじ引き開始！（１人目)"

IF password <> 1 Then do while (MsgBox(msg,vbYesNo) = vbYes)

    rand_num = Int(Rnd * user_num)

    find_num(rand_num) = 1

    IF find_num(rand_num) = intwinNumbers(rand_num) Then
        MsgBox "当たり！"
        win_count = win_count + 1
        If win_num_2 = win_count Then exit do
        intwinNumbers(rand_num) = 0
        user_num = user_num - 1
        For intCounter_5 = rand_num To user_num
            intwinNumbers(intCounter_5) = intwinNumbers(intCounter_5 + 1)
        Next
    else
        MsgBox "はずれ"
        user_num = user_num - 1
        For intCounter_5 = rand_num To user_num
            intwinNumbers(intCounter_5) = intwinNumbers(intCounter_5 + 1)
        Next
    End If

    intLoop_cnt = intLoop_cnt + 1
    msg = "くじ引き開始！("&intLoop_cnt&"人目)"
loop
else do while (MsgBox(msg,vbYesNo) = vbYes)

        If  intLoop_cnt <> debag_win Then
            MsgBox "はずれ"
        else
            MsgBox "当たり！" 
            exit do
        End IF 

        intLoop_cnt = intLoop_cnt + 1
        msg = "くじ引き開始！("&intLoop_cnt&"人目)"
    loop
End IF