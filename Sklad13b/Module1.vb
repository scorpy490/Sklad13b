Module Module1

    Sub Main()
        Dim ConnSQL, ts1, sqlstr, fl, path, buf, arr, folder, kup, kdef, d, sqlstr1, errfl, k, rs1
        ConnSQL = CreateObject("ADODB.Connection")
        ConnSQL.ConnectionString = "Provider=SQLOLEDB;Server=srv-otk;Database=otk;Trusted_Connection=yes;Integrated Security=SSPI;Persist Security Info=False"
        Dim fso, fserr, goreem
        Dim sort, sort2 As String
        Dim dbupd(10000) As String
        Dim gormc = 0

        ConnSQL.Open
        fso = CreateObject("Scripting.FileSystemObject")
        fserr = CreateObject("Scripting.FileSystemObject")

        errfl = fserr.OpenTextFile("c:\Terminal\err\err.txt", 8, True)
        fl = ""
        path = ""
        folder = fso.GetFolder("C:\Terminal\OUT")
        kup = 0
        kdef = 0
        d = 0
        k = 0
        For Each file In Folder.Files
            If Left(file.Name, 3) = "skl" And Right(file.name, 4) = ".txt" Then
                path = file
                fl = file.name
                Exit For
            End If
        Next
        If Len(fl) < 5 Then
            MsgBox("Файл данных не найден!", vbOKCancel, "Ошибка")

            'System.Threading.Thread.Sleep(7000)
            Exit Sub
            End If



            ts1 = fso.OpenTextFile(path, 1, False)
            Do While Not ts1.AtEndOfStream
                buf = ts1.ReadLine
            arr = Split(buf, ";")
            sqlstr1 = "SELECT [shtr_kod], [Сорт],[sort13] FROM dbo.[Изделия] WHERE [shtr_kod]=" & arr(2)
            rs1 = ConnSQL.Execute(sqlstr1)
            If rs1.EOF = True Then

                Console.WriteLine(arr(2) & " не существует")
                errfl.WriteLine(CDate(arr(0) & " " & arr(1)) & vbTab & Now.ToShortTimeString & vbTab & arr(2) & " не существует")
                d = d + 1
                Continue Do
            End If
            sort = rs1(1).Value.ToString
            If sort = "2" Or sort = "6" Or sort = "7" Then sort = "1"
            sort2 = rs1(2).Value.ToString
            If sort2 = "" Then sort2 = sort

            If Left(arr(3), 1) = 2 Then
                sqlstr = "Update dbo.Изделия SET [DataUp] ='" & CDate(arr(0) & " " & arr(1)) & "', [NomUp] =" & Mid(arr(3), 2, 11) & ", [Sort13]=" & sort2 & "WHERE [shtr_kod]=" & arr(2)
                kup = kup + 1

            Else
                If arr(4) = 9 Then
                    arr(4) = 4
                    goreem = 2
                    gormc = gormc + 1
                Else
                    goreem = 0
                End If

                If arr(4) = "" Then
                    sort2 = "4"
                Else
                    sort2 = arr(4)
                End If
                'sqlstr = "INSERT INTO dbo.t1 (d1,shtr,upak, def) SELECT '" & Cdate (arr(0) &" "&arr(1)) &"'," &arr(2) &", null," & arr(3)
                sqlstr = "Update dbo.Изделия SET [DataUp] ='" & CDate(arr(0) & " " & arr(1)) & "', DefUp =" & arr(3) & ",[sort13]=" & sort2 & ", goreem=" & goreem & ",[pereat]=1, [NomUp]=null WHERE [shtr_kod]=" & arr(2)
                kdef = kdef + 1

                End If
            'ConnSQL.execute = sqlstr
            dbupd(k) = sqlstr
            k = k + 1
            ConnSQL.execute("Update dbo.sklad SET [13skl]='1' WHERE [shtr]=" & arr(2))

        Loop
            sqlstr = "INSERT INTO dbo.LogUpak (data,CountUp,CountBrak) SELECT getdate()," & kup & "," & kdef
        'ConnSQL.execute = sqlstr


        ConnSQL.Close


            ConnSQL.Open
        ConnSQL.BeginTrans
        For i = 0 To k - 1
            ConnSQL.Execute(dbupd(i))
            'MsgBox(dbupd(i))
        Next
        ConnSQL.CommitTrans
        ConnSQL.Close
        ts1.Close
        errfl.Close
        ts1 = fso.GetFile(path)
        Try
            ts1.move("C:\terminal\Arhiv\" & fl)
        Catch ex As Exception
            ts1.delete("C:\terminal\OUT\" & fl)
        End Try

        MsgBox("Упаковано: " & kup & vbNewLine & "Переатестаций: " & kdef & vbNewLine & "Не найдено: " & d & vbNewLine & "На реэмалирование:" & gormc, vbOKOnly, "Данные успешно загружены")
        'Console.WriteLine("Переатестаций: " & kdef)
        'Console.WriteLine("Не найдено: " & d)

        'System.Threading.Thread.Sleep(7000)
    End Sub

End Module
