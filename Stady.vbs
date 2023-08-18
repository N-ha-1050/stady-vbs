' �����C�N 2023-8-18

option explicit
randomize


dim so, sa, ws, wa
set so = CreateObject("Scripting.FileSystemObject")
set sa = WScript.CreateObject("Shell.Application")
set ws = CreateObject("WScript.Shell")
set wa = WScript.Arguments


dim inipath, csvpath, q, a, msg, record

if wa.Count = 1 then

    inipath = wa(0)

    dim inifile
    set inifile = so.getfile(inipath)

    ws.CurrentDirectory = inifile.parentfolder

else

    dim gf
    set gf = so.getfolder(".")

    dim fileselectcnt
    fileselectcnt = 0

    dim filepath, fileselectmsg, inifiles()
    for each filepath in gf.files

        dim file
        set file = so.OpenTextFile(filepath)
        
        do until file.AtEndOfStream

            dim fileread
            fileread = file.readline

            if ucase(trim(fileread)) = "[MEMORIES QUESTION]" then
                fileselectcnt = fileselectcnt + 1
                fileselectmsg = fileselectmsg & fileselectcnt & ". " & so.getbasename(filepath) & vbcrlf

                redim Preserve inifiles(fileselectcnt)
                inifiles(fileselectcnt) = filepath
            end if
        loop
        file.close
    next

    dim ininum
    ininum = inputbox("�ݒ�t�@�C����I�����Ă��������B" & vbcrlf & fileselectmsg, , 1)

    if ininum = "" then
        Wscript.Quit
    end if

    inipath = inifiles(ininum)

end if

dim ini
set ini = so.OpenTextFile(inipath)

dim sec
sec = ""

do until ini.AtEndOfStream

    dim iniread
    iniread = ini.readline

    if not(trim(iniread) = "") and not(mid(trim(iniread), 1, 1) = "#") then

        if mid(trim(iniread), 1, 1) = "[" then
            sec = ucase(trim(iniread))

        elseif sec = "[MEMORIES QUESTION]" then

            dim inient
            inient = split(iniread, "=")

            if ubound(inient) = 1 then
                select case ucase(trim(inient(0)))

                    case "CSV"
                        csvpath = trim(inient(1))
                    case "Q"
                        q = trim(inient(1))
                    case "A"
                        a = trim(inient(1))
                    case "MSG"
                        msg = trim(inient(1))
                    case "DATA"
                        record = trim(inient(1))
                    case "RECORD"
                        record = trim(inient(1))


                end select
            end if
        end if
    end if
loop

ini.close

dim linecsv
set linecsv = so.OpenTextFile(csvpath)

dim linenum
linenum = 0

do until linecsv.AtEndOfStream

    dim linecsvread
    linecsvread = linecsv.readline

    if not(trim(linecsvread) = "") and not(mid(trim(linecsvread), 1, 1) = "#") and not(mid(trim(linecsvread), 1, 2) = "//") then
        linenum = linenum + 1
    end if
loop

linecsv.close

dim csv
set csv = so.OpenTextFile(csvpath)

dim data()

dim cnt
cnt = 0

do until csv.AtEndOfStream

    dim csvread
    csvread = csv.readline

    if not(trim(csvread) = "") and not(mid(trim(csvread), 1, 1) = "#") and not(mid(trim(csvread), 1, 2) = "//") then
        
        dim readdata
        readdata = split(csvread, ",")

        redim Preserve data(linenum - 1, ubound(readdata))

        dim i
        for i = 0 to ubound(readdata)
            data(cnt, i) = trim(readdata(i))
        next

        cnt = cnt + 1
    end if
loop

csv.Close

if record = 1 then

    dim recordcsvpath
    recordcsvpath = inipath & ".csv"

    if not(so.fileexists(recordcsvpath)) then

        dim recordnewcsv
        set recordnewcsv = so.createtextfile(recordcsvpath)

        recordnewcsv.writeline("Memories Question Data")
        ' recordnewcsv.writeline(q)
        recordnewcsv.writeline("���, ����, �o�葍��, ������, �듚")

        dim j
        for j = 1 to linenum - 1
            recordnewcsv.writeline(data(j, q) & ", " & data(j, a) & ", 0, 0, ")
        next

        recordnewcsv.close

        dim recordnewcsvfile
        set recordnewcsvfile = so.getfile(recordcsvpath)
        recordnewcsvfile.attributes = 2

        msgbox("�L�^�f�[�^��V�K�쐬���܂����B")
    end if

    dim recordcsv
    set recordcsv = so.OpenTextFile(recordcsvpath)

    dim DataQuestion(), DataAnswer(), DataAll(), DataRight(), DataWrong()
    redim DataQuestion(linenum - 1), DataAnswer(linenum - 1), DataAll(linenum - 1), DataRight(linenum - 1), DataWrong(linenum - 1)

    dim dff
    dff = ucase(trim(recordcsv.readline))

    if not(dff = "MEMORIES QUESTION DATA") then

        recordcsv.close

        dim dffdel
        dffdel = msgbox("data�t�@�C�����j�����Ă��܂��B" & vbcrlf & "data �t�@�C�����폜���܂����H", 276)

        if dffdel = 6 then
            so.deletefile(recordcsvpath)
        end if

        wscript.quit
    end if

    dim recordcnt
    recordcnt = 0

    do until recordcsv.AtEndOfStream

        dim recordcsvread
        recordcsvread = recordcsv.readline

        if not(recordcsvread = "") then

            dim recordcsvdata
            recordcsvdata = split(recordcsvread, ",")

            DataQuestion(recordcnt) = trim(recordcsvdata(0))
            DataAnswer(recordcnt) = trim(recordcsvdata(1))
            DataAll(recordcnt) = trim(recordcsvdata(2))
            DataRight(recordcnt) = trim(recordcsvdata(3))
            DataWrong(recordcnt) = trim(recordcsvdata(4))

            if recordcnt > 0 and not(DataQuestion(recordcnt) = data(recordcnt, q) and DataAnswer(recordcnt) = data(recordcnt, a)) then

                recordcsv.close
            
                dim linedel
                linedel = msgbox("data�t�@�C�����j�����Ă��܂��B" & vbcrlf & "data �t�@�C�����폜���܂����H", 276)
            
                if linedel = 6 then
                    so.deletefile(recordcsvpath)
                end if
            
                wscript.quit
            end if

            recordcnt = recordcnt + 1
        end if
    loop
    recordcsv.close
end if

do
    dim num
    num = int(rnd * (linenum - 1)) + 1

    dim ans
    ans = inputbox(msg & vbcrlf & data(num, q))
    
    if isempty(ans) then

        if record = 1 then

            dim recordwritecsv
            set recordwritecsv = so.OpenTextFile(recordcsvpath, 2)
            
            recordwritecsv.writeline("Memories Question Data")
            recordwritecsv.writeline("���, ����, �o�葍��, ������, �듚")

            dim k
            for k = 1 to linenum - 1
                

                if not(DataQuestion(k) = data(k, q) and DataAnswer(k) = data(k, a)) then

                    recordwritecsv.close
                
                    dim writedel
                    writedel = msgbox("data�t�@�C�����j�����Ă��܂��B" & vbcrlf & "data �t�@�C�����폜���܂����H", 276)
                
                    if writedel = 6 then
                        so.deletefile(recordcsvpath)
                    end if
                
                    wscript.quit
                end if
                recordwritecsv.writeline(DataQuestion(k) & ", " & DataAnswer(k) & ", " & DataAll(k) & ", " & DataRight(k) & ", " & DataWrong(k))
            next

            recordwritecsv.close

            msgbox("�Z�[�u���܂����B")
        end if
        wscript.quit
    
    elseif ans = data(num, a) then
        if record = 1 then
            DataAll(num) = int(DataAll(num)) + 1
            DataRight(num) = int(DataRight(num)) + 1
        end if

    else
        if record = 1 then
            DataAll(num) = int(DataAll(num)) + 1
            DataWrong(num) = DataWrong(num) & ans & "; "
        end if
        msgbox("�s�����ł��B" & vbcrlf & "���F" & data(num, q) & vbcrlf & "�듚: " & ans & vbcrlf & "����: " & data(num, a))
    end if
loop