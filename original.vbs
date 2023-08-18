'自動ファイル読み込み・記録機能対応用
Option Explicit
Dim txt,readtxt,txtent,ini,iniread,csvpath,so,dn,dnn,di,wa,dw,io,coun,gi,dread,dff,sa,cv,cou,iw,iwi,datacsvpath,inipath,gf,data,fileselctmsg,fil,f,i,count,cvread,line,cv2,Qmsg,Q,Amsg,A,msg,num,ans,i1,i2,i3,c,csv(),dsplit,DataQue(),DataAns(),DataAll(),DataRig(),DataNot(),testi,testmsg
Randomize
Set so = CreateObject("Scripting.FileSystemObject")
Set sa = WScript.CreateObject("Shell.Application")
Set wa = WScript.Arguments

If wa.Count = 1 Then

Set txt = so.OpenTextFile(wa(0))

inipath = wa(0)

Do Until txt.AtEndOfStream
readtxt = txt.ReadLine

If Not(trim(readtxt)="") Then
If Not(Mid(trim(readtxt),1,1)="#") Then
    if Ucase(Trim(readtxt)) = "[" then
        sec = Ucase(Trim(readtxt))
    else
        txtent = Split(readtxt,"=")
        if ( Ubound(txtent) = 1 ) then
            if Ucase(Trim(txtent(0)))="CSV" then
                csvpath = txtent(1)
            end if
            if Ucase(Trim(txtent(0)))="Q" then
                Q = txtent(1)
            end if
            if Ucase(Trim(txtent(0)))="A" then
                A = txtent(1)
            end if
            if Ucase(Trim(txtent(0)))="MSG" then
                msg = txtent(1)
            end if
            if Ucase(Trim(txtent(0)))="DATA" then
                data = txtent(1)
            end if
        end if
    end if
end if
end if

Loop

txt.close

Else

set gf = so.getfolder(".")

Dim file()
cou = 0
For Each c In gf.Files
Set ini = so.OpenTextFile(c.name,1)
do
iniread = ini.readline
loop while iniread = ""
if Ucase(Trim(iniread)) = "[MEMORIES QUESTION]" then
cou = cou + 1
fileselctmsg = fileselctmsg & cou & "." & c.Name & vbCrLf
ReDim Preserve file(cou)
file(cou) = c.name
End If
Next
ini.close
fil = inputbox("設定ファイルを選択してください。" & vbCrLf & fileselctmsg)
If fil = "" Then
Wscript.Quit
end if

inipath = file(fil)

Set txt = so.OpenTextFile(file(fil),1)

Do Until txt.AtEndOfStream
readtxt = txt.ReadLine

If Not(trim(readtxt)="") Then
If Not(Mid(trim(readtxt),1,1)="#") Then
    if Ucase(Trim(readtxt)) = "[" then
        sec = Ucase(Trim(readtxt))
    else
        txtent = Split(readtxt,"=")
        if ( Ubound(txtent) = 1 ) then
            if Ucase(Trim(txtent(0)))="CSV" then
                csvpath = txtent(1)
            end if
            if Ucase(Trim(txtent(0)))="Q" then
                Q = txtent(1)
            end if
            if Ucase(Trim(txtent(0)))="A" then
                A = txtent(1)
            end if
            if Ucase(Trim(txtent(0)))="MSG" then
                msg = txtent(1)
            end if
            if Ucase(Trim(txtent(0)))="DATA" then
                data = txtent(1)
            end if
        end if
    end if
end if
end if

Loop

End If
txt.close

Set cv2 = so.OpenTextFile(csvpath,8)
line = cv2.line
cv2.close

Set cv = so.OpenTextFile(csvpath)

count = 0
Do Until cv.AtEndOfStream
cvread=cv.ReadLine
If Not(cvread = "") then
If Not(Mid(trim(cvread),1,2)="//") Then
c = Split(cvread, ",")
ReDim Preserve csv(line,UBound(c))
for i = 0 to UBound(c)
csv(count,i)=c(i)
next

count = count+1
End If
End If
Loop
cv.Close


datacsvpath=inipath&".csv"
Set dn = so.OpenTextFile(datacsvpath,1,true)
dn.close

Set gi = so.GetFile(inipath&".csv")
gi.attributes = 2
If FormatNumber(gi.Size, 0) = 0 then
Set dnn = so.OpenTextFile(datacsvpath,2)
dnn.writeline("new file")
dnn.close
msgbox "記録データを新規作成しました。"
End If


Set io = so.OpenTextFile(datacsvpath,1,true)
dff = Ucase(Trim(io.ReadLine))
io.close
If Not(dff = "MEMORIES QUESTION DATA") Then

Set so = WScript.CreateObject("Scripting.FileSystemObject")

Set iw = so.OpenTextFile(datacsvpath,2)
iw.WriteLine("Memories Question Data")
iw.WriteLine("問題,正答,出題総数,正解回数,誤答")
for iwi = 1 to count
iw.WriteLine(csv(iwi,Q)&","&csv(iwi,A)&",0,0,")
next
iw.close
End If
Set io = so.OpenTextFile(datacsvpath)
dff = Ucase(Trim(io.ReadLine))

coun=0
Do Until io.AtEndOfStream
dread = io.readline
If Not(dread = "") then
ReDim Preserve DataQue(coun)
ReDim Preserve DataAns(coun)
ReDim Preserve DataAll(coun)
ReDim Preserve DataRig(coun)
ReDim Preserve DataNot(coun)
dsplit = Split(dread, ",")
DataQue(coun)=dsplit(0)
DataAns(coun)=dsplit(1)
DataAll(coun)=dsplit(2)
DataRig(coun)=dsplit(3)
DataNot(coun)=dsplit(4)

coun=coun+1
end If
Loop



io.close



Do
num = Int(Rnd * (count-1))+1
ans = InputBox(msg&vbcrlf&csv(num,Q))
If Not(DataQue(num)=csv(num,Q)) or Not(DataAns(num)=csv(num,A)) Then
msgbox "dataファイルが破損しています。" & vbcrlf & DataQue(num)&"="&csv(num,Q)&vbcrlf&DataAns(num)&"="&csv(num,A)
WScript.Quit


End If
If IsEmpty(ans) Then
Set dw = so.OpenTextFile(datacsvpath,2)
dw.WriteLine("Memories Question Data")
for di = 0 to UBound(DataQue)
dw.WriteLine(DataQue(di)&","&DataAns(di)&","&DataAll(di)&","&DataRig(di)&","&DataNot(di))
next
dw.close
msgbox "セーブしました。"
WScript.Quit
ElseIf ans = csv(num,A) Then
DataAll(num)=Int(DataAll(num))+1
DataRig(num)=Int(DataRig(num))+1
'MsgBox "正解です。"&vbcrlf&csv(num,Q)&"=>"&csv(num,A)
Else
DataAll(num)=Int(DataAll(num))+1
'DataNot(num)=DataNot(num) & space(1) & ans
DataNot(num)=DataNot(num) & "'" & ans & "'"

MsgBox "不正解です。"&vbcrlf&"問："&csv(num,Q)&vbcrlf&"誤："&ans&vbcrlf&"正："&csv(num,A)
End If
Loop