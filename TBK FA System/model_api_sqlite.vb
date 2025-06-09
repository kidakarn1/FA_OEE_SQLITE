Imports System.Globalization
Imports System.Web.Script.Serialization
Imports Newtonsoft.Json

Public Class model_api_sqlite
    Public Shared Async Function UpdateStatus_tag_print_detail() As Task(Of String)
        If ins_qty.Visible = False Then
            Dim date_st = DateTime.Now.ToString("yyyy-MM-dd") & " 00:00:00"
            Dim date_end = DateTime.Now.ToString("yyyy-MM-dd") & " 23:59:59"
            Dim Sql = "Select * from tag_print_detail where tr_status = '0'"
            Dim api = New api
            Dim jsonData As String = Await api.Load_dataSQLiteAsyncLoaddata(Sql)
            Try
                Dim i As Integer = 0
                Dim dcResultdata As Object = New JavaScriptSerializer().Deserialize(Of List(Of Object))(jsonData)
                For Each items As Object In dcResultdata
                    Dim id = items("id").ToString()
                    Dim wi = items("wi").ToString()
                    Dim qr_detail = items("qr_detail").ToString()
                    Dim box_no = items("box_no").ToString()
                    Dim print_count = items("print_count").ToString()
                    Dim seq_no = items("seq_no").ToString()
                    Dim shift = items("shift").ToString()
                    Dim flg_control = items("flg_control").ToString()
                    Dim item_cd As String = qr_detail.Substring(20, 35).Trim()
                    Dim pwi_id = items("pwi_id").ToString()
                    Dim tag_group_no = items("tag_group_no").ToString()
                    Dim nextProcess = items("next_proc").ToString()
                    Dim tr_status = 1
                    Dim pk_tag_print_id = Await Backoffice_model.Trasnfer_tag_print_detail(wi, qr_detail, box_no, print_count, seq_no, shift, flg_control, item_cd, pwi_id, tag_group_no, "0", nextProcess, tr_status)
                    If pk_tag_print_id <> 0 Then
                        Dim sqlUpdate = $"Update tag_print_detail set tr_status = '{tr_status}' where id = '{id}'"
                        Dim jsonDataUpdate As String = Await api.Load_dataSQLiteAsync(sqlUpdate)
                    End If
                    i += 1
                    If i = dcResultdata.Count Then ' ช่วยลด การทำงานหนัก จากปกติต้องทำงาน ทุุกครั้งรอ ตอนจบทีเดียว ค่อยทำ 
                        Await model_api_sqlite.UpdateStatus_tag_print_detail_sub(pk_tag_print_id, wi)
                        Await model_api_sqlite.UpdateStatus_tag_print_detail_main()
                    End If
                Next
            Catch ex As Exception
            End Try
        End If
        Return "Done"
    End Function



    Public Shared Async Function UpdateStatus_tag_print_detail_main() As Task(Of String)
        Dim date_st = DateTime.Now.ToString("yyyy-MM-dd") & " 00:00:00"
        Dim date_end = DateTime.Now.ToString("yyyy-MM-dd") & " 23:59:59"
        'Dim Sql = "Select * from tag_print_detail_main where created_date between '" & date_st & "' and  '" & date_end & "' and tr_status = '0'"
        Dim Sql = "Select * from tag_print_detail_main where  tr_status = '0'"
        'Console.WriteLine(Sql)
        Dim api = New api
        Dim jsonData As String = Await api.Load_dataSQLiteAsyncLoaddata(Sql)
        '  Try
        Dim dcResultdata As Object = New JavaScriptSerializer().Deserialize(Of List(Of Object))(jsonData)
        For Each items As Object In dcResultdata
            Dim tag_ref_str_id = items("tag_ref_str_id").ToString()
            Dim tag_ref_end_id = items("tag_ref_end_id").ToString()
            Dim line_cd As String = items("line_cd").ToString()
            Dim tag_qr_detail = items("tag_qr_detail").ToString()
            Dim seq_no = items("tag_qr_detail").ToString().Substring(95, 3)
            Dim tag_batch_no = items("tag_batch_no").ToString()
            Dim tag_next_proc = items("tag_next_proc").ToString()
            Dim flg_control = items("flg_control").ToString()
            Dim created_date = items("created_date").ToString()
            Dim updated_date = items("updated_date").ToString()
            Dim tag_wi_no = items("tag_wi_no").ToString()
            Dim pwi_no = items("pwi_id").ToString()
            Dim tag_group_no = items("tag_group_no").ToString()
            Dim id = items("tag_id").ToString()
            Dim lot_no = items("lot_no").ToString()
            Dim tr_status = 1
            Dim pkMainId = Await Backoffice_model.Trasnfer_tag_print_main(tag_ref_str_id, tag_ref_end_id, line_cd, tag_qr_detail, tag_batch_no, tag_next_proc, flg_control, created_date, updated_date, tag_wi_no, pwi_no, tag_group_no, tr_status, seq_no, lot_no)
            If pkMainId <> 0 Then
                Dim sqlUpdate = "Update tag_print_detail_main set tr_status = '" & tr_status & "' where tag_id = '" & id & "'"
                Console.WriteLine("sqlUpdate tag print detail main===>" & sqlUpdate)
                Dim jsonDataUpdate As String = Await api.Load_dataSQLiteAsync(sqlUpdate)
            End If
        Next
        'Catch ex As Exception
        ' End Try
    End Function
    Public Shared Async Function UpdateStatus_tag_print_detail_sub(tag_print_detail_id As String, wi As String) As Task(Of String)
        Dim date_st = DateTime.Now.ToString("yyyy-MM-dd") & " 00:00:00"
        Dim date_end = DateTime.Now.ToString("yyyy-MM-dd") & " 23:59:59"
        '  Dim Sql = "Select * from tag_print_detail_sub where created_date between '" & date_st & "' and  '" & date_end & "' and tr_status = '0'"
        Dim Sql = "Select * from tag_print_detail_sub where   tr_status = '0'"
        'Console.WriteLine(Sql)
        Dim api = New api
        Dim jsonData As String = Await api.Load_dataSQLiteAsyncLoaddata(Sql)
        Try
            Dim dcResultdata As Object = New JavaScriptSerializer().Deserialize(Of List(Of Object))(jsonData)
            For Each items As Object In dcResultdata
                Dim tag_id = items("tag_id").ToString()
                Dim line_cd = items("line_cd").ToString()
                Dim tag_qr_detail = items("tag_qr_detail").ToString()
                Dim flg_control = items("flg_control").ToString()
                Dim created_date = items("created_date").ToString()
                Dim updated_date = items("updated_date").ToString()
                Dim tag_wi_no = items("tag_wi_no").ToString()
                Dim lot_no = items("lot_no").ToString()
                Dim tr_status = 1
                Await Backoffice_model.Transfer_Tag_Print_sub(wi, tag_print_detail_id, line_cd, tag_qr_detail, flg_control, created_date, updated_date, tag_wi_no, "1")
                Dim sqlUpdate = "Update tag_print_detail_sub set tr_status = '" & tr_status & "' where tag_id = '" & tag_id & "'"
                ' Console.WriteLine(sqlUpdate)
                Dim jsonDataUpdate As String = Await api.Load_dataSQLiteAsync(sqlUpdate)
            Next
        Catch ex As Exception
        End Try
    End Function
    Public Shared Async Function mas_INSERT_production_working_info(ind_row As String, pwi_lot_no As String, pwi_seq_no As String, pwi_shift As String, pwi_id As String) As Task(Of String)
        Dim dateTime_Crr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Try
            Dim Sql = "Insert into production_working_info(
            pwi_id , 
			ind_row , 
			pwi_lot_no , 
			pwi_seq_no , 
			pwi_shift , 
			pwi_created_date , 
			pwi_created_by) 
			Values(
                '" & pwi_id & "' , 
				'" & ind_row & "' , 
				'" & pwi_lot_no & "' , 
				'" & pwi_seq_no & "' , 
				'" & pwi_shift & "' , 
				'" & dateTime_Crr & "' , 
				'SYSTEM')"
            '   Console.WriteLine(Sql)
            Dim api = New api
            Dim jsonData As String = Await api.Load_dataSQLiteAsync(Sql)
            Return 1
        Catch ex As Exception
            MsgBox("Error Files model_api_sqlite In Function mas_INSERT_production_working_info")
        End Try
    End Function
    Public Shared Async Function mas_Insert_tag_print(wi As String, qr_detail As String, box_no As Integer, print_count As Integer, seq_no As String, shift As String, flg_control As Integer, item_cd As String, pwi_id As String, tag_group_no As String, goodQty As Integer, nextProcess As String, tr_status As Integer) As Task(Of String)
        Try
            Dim currdated As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Await mas_update_tagprint(wi, "2", "0", tr_status)
            tag_group_no = "1"
            Dim Sql = "INSERT INTO tag_print_detail(wi,qr_detail,box_no,print_count,created_date,updated_date,seq_no,shift , next_proc ,  flg_control , pwi_id , tag_group_no , tr_status) VALUES ('" & wi & "','" & qr_detail & "','" & box_no & "','" & print_count & "','" & currdated & "','" & currdated & "','" & seq_no & "','" & shift & "','" & nextProcess & "' ,'" & flg_control & "','" & pwi_id & "','" & tag_group_no & "','" & tr_status & "')"
            'Console.WriteLine(Sql)
            Dim api = New api
            Dim jsonData As String = Await api.Load_dataSQLiteAsync(Sql)
            Return 1
        Catch ex As Exception
            MsgBox("Error Files model_api_sqlite In Function mas_Insert_tag_print")
        End Try
    End Function
    Public Shared Async Function mas_get_tag_print_detail_main(qr_detail As String) As Task(Of String)
        Try
            Dim api = New api
            Dim Sql = " Select  tag_id  As id_print from tag_print_detail_main where tag_qr_detail = '" & qr_detail & "'"
            'Console.WriteLine(Sql)
            Dim jsonData As String = Await api.Load_dataSQLiteAsyncLoaddata(Sql)
            Dim id_print As String = 0
            Try
                Dim dcResultdata As Object = New JavaScriptSerializer().Deserialize(Of List(Of Object))(jsonData)
                For Each items As Object In dcResultdata
                    id_print = items("id_print").ToString()
                Next
            Catch ex As Exception

            End Try
            Return id_print
        Catch ex As Exception
            MsgBox("Error Files model_api_sqlite In Function mas_get_tag_print_detail_main")
        End Try
    End Function
    Public Shared Async Function mas_Insert_tag_print_sub(ref_id As String, line As String, qr_code As String, wi As String, tag_group_no As String, next_process As String, tr_status As Integer, lot_no As String) As Task(Of String)
        Try
            Dim currdated As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Await mas_update_tagprint_sub(wi, "2", "0", tr_status)
            Dim tmp_tag_group_no As Integer = 1
            Dim Sql = "INSERT INTO tag_print_detail_sub(tag_ref_id , line_cd , tag_qr_detail , flg_control , created_date , updated_date , tag_wi_no , tag_group_no , tr_status , lot_no) VALUES ('" & ref_id & "','" & line & "','" & qr_code & "' ,'" & print_back.check_tagprint_main() & "' , '" & currdated & "' , '" & currdated & "' , '" & wi & "' , '" & tmp_tag_group_no & "' , '" & tr_status & "' , '" & lot_no & "')"
            Console.WriteLine(Sql)
            Dim api = New api
            Dim jsonData As String = Await api.Load_dataSQLiteAsync(Sql)
            Return 1
        Catch ex As Exception
            MsgBox("Error Files model_api_sqlite In Function mas_Insert_tag_print_sub")
        End Try
    End Function
    Public Shared Async Function mas_Insert_tag_print_main(wi As String, qr_detail As String, batch_no As Integer, print_count As Integer, seq_no As String, shift As String, flg_control As Integer, item_cd As String, pwi_id As String, tag_group_no As String, next_process As String, tr_status As Integer, lot_no As String) As Task(Of String)
        Try
            ' MsgBox("ready insert Main model")
            Dim currdated As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Await mas_update_tagprint(wi, "2", "0", tr_status)
            Await mas_update_tagprint_main(wi, "2", "0", tr_status)
            Dim start_id As String = Await mas_Get_ref_start_id(wi, seq_no, Working_Pro.Label18.Text)
            Dim end_id As String = Await mas_Get_ref_end_id(wi, seq_no, Working_Pro.Label18.Text)
            tag_group_no = "1"
            Dim Sql = "INSERT INTO tag_print_detail_main(tag_ref_str_id ,tag_ref_end_id , line_cd , tag_qr_detail , tag_batch_no , tag_next_proc , flg_control , created_date , updated_date , tag_wi_no , pwi_id , tag_group_no , tr_status , lot_no) VALUES ('" & start_id & "','" & end_id & "','" & MainFrm.Label4.Text & "','" & qr_detail & "' ,'" & batch_no & "' ,'" & next_process & "','" & flg_control & "','" & currdated & "','" & currdated & "','" & wi & "','" & pwi_id & "' ,'" & tag_group_no & "' ,'" & tr_status & "','" & lot_no & "')"
            Console.WriteLine(Sql)
            Dim api = New api
            Dim jsonData As String = Await api.Load_dataSQLiteAsync(Sql)
            Return 1
        Catch ex As Exception
            MsgBox("Error Files model_api_sqlite In Function mas_Insert_tag_print_main")
        End Try
    End Function
    Public Shared Async Function mas_Get_ref_start_id(wi As String, seq_no As String, lot_no As String) As Task(Of String)
        Dim api = New api()
        Dim Sql = " select min(id) as id from tag_print_detail where wi = '" & wi & "'  and  SUBSTRING( qr_detail, 96, 3) = '" & seq_no & "' and SUBSTRING( qr_detail, 59, 4) = '" & lot_no & "'"
        'Console.WriteLine(Sql)
        Dim id As Integer = 0
        Dim jsonData As String = Await api.Load_dataSQLiteAsyncLoaddata(Sql)
        Try
            Dim dcResultdata As Object = New JavaScriptSerializer().Deserialize(Of List(Of Object))(jsonData)
            For Each items As Object In dcResultdata
                id = items("id").ToString()
            Next
        Catch ex As Exception
        End Try
        Return id
    End Function
    Public Shared Async Function mas_Get_ref_end_id(wi As String, seq_no As String, lot_no As String) As Task(Of String)
        Dim api = New api()
        Dim Sql = " select max(id) as id from tag_print_detail where wi = '" & wi & "'  and  SUBSTRING( qr_detail, 96, 3) = '" & seq_no & "' and SUBSTRING( qr_detail, 59, 4) = '" & lot_no & "'"
        'Console.WriteLine(Sql)
        Dim id As Integer = 0
        Dim jsonData As String = Await api.Load_dataSQLiteAsyncLoaddata(Sql)
        Try
            Dim dcResultdata As Object = New JavaScriptSerializer().Deserialize(Of List(Of Object))(jsonData)
            For Each items As Object In dcResultdata
                id = items("id").ToString()
            Next
        Catch ex As Exception
        End Try
        Return id
    End Function


    Public Shared Async Function mas_update_tagprint(wi As String, flgUpdate As String, conditionflg As String, tr_status As Integer) As Task(Of String)        '2 , 0
        Try
            Dim currdated As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Dim Sql = "update tag_print_detail set flg_control = '" & flgUpdate & "' , tr_status ='" & tr_status & "' where flg_control = '" & conditionflg & "' and  wi = '" & wi & "'"
            'Console.WriteLine(Sql)
            Dim api = New api
            Dim jsonData As String = Await api.Load_dataSQLiteAsync(Sql)
            Return 1
        Catch ex As Exception
            MsgBox("Error Files model_api_sqlite In Function mas_update_tagprint")
        End Try
    End Function
    Public Shared Async Function mas_update_tagprint_main(wi As String, flgUpdate As String, conditionflg As String, tr_status As Integer) As Task(Of String)        '2 , 0
        Try
            Dim currdated As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Dim Sql = "update tag_print_detail_main set flg_control = '" & flgUpdate & "' , tr_status ='" & tr_status & "' where flg_control = '" & conditionflg & "' and  tag_wi_no = '" & wi & "'"
            ' Console.WriteLine(Sql)
            Dim api = New api
            Dim jsonData As String = Await api.Load_dataSQLiteAsync(Sql)
            Return 1
        Catch ex As Exception
            MsgBox("Error Files model_api_sqlite In Function mas_update_tagprint_main")
        End Try
    End Function
    Public Shared Async Function mas_update_tagprint_sub(wi As String, flgUpdate As String, conditionflg As String, tr_status As Integer) As Task(Of String)        '2 , 0
        Try
            Dim currdated As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Dim Sql = "update tag_print_detail set flg_control = '" & flgUpdate & "' , tr_status ='" & tr_status & "' where flg_control = '" & conditionflg & "' and  wi = '" & wi & "'"
            'Console.WriteLine(Sql)
            Dim api = New api
            Dim jsonData As String = Await api.Load_dataSQLiteAsync(Sql)
            Return 1
        Catch ex As Exception
            MsgBox("Error Files model_api_sqlite In Function mas_update_tagprint")
        End Try
    End Function


    Public Shared Async Function mas_Update_seqplan(wi As String, line_cd As String, date_start As String, date_end As String, Update_seq As String) As Task(Of String)
        Dim api = New api
        Try
            Dim currdated As String = DateTime.Now.ToString("yyyy/MM/dd")
            Dim today As Date = Date.Today
            Dim time_tomorrow As DateTime = today.AddDays(1)
            Dim format_tommorow = "yyyy/MM/dd"
            Dim date_tommorow = time_tomorrow.ToString(format_tommorow)
            date_end_covert = date_tommorow & " 07:59:59"
            Try
                Dim time_now As DateTime
                time_now = DateTime.Now.ToString("hh:mm:ss tt")
                If time_now >= "08:00:00 AM" And time_now <= "07:59:59 PM" Then
                    date_start = currdated & " 08:00:00"
                    ' date_start = date_start & " 08:00:00"
                Else
                    date_start = currdated & " 08:00:00"
                End If
                If time_now >= "12:00:00 AM" And time_now <= "08:00:00 AM" Then
                    Dim format_tommorow_re = "yyyy/MM/dd"
                    Dim del_date1 As DateTime = today.AddDays(-1)
                    date_start = del_date1.ToString(format_tommorow_re)
                    Dim sub_date_end1 = Trim(date_end.ToString.Substring(0, 10))
                    date_start = date_start & " 08:00:00"
                    date_end_covert = sub_date_end1 & " 07:59:59"
                End If
            Catch ex As Exception

            End Try
            Try
                Dim sql = "SELECT tmp_id from tmp_planseq where tmp_created_date BETWEEN  '" & date_start & "' and '" & date_end_covert & "' and tmp_line_cd = '" & line_cd & "'"
                'Console.WriteLine("sql====>" & sql)
                Dim jsonData As String = Await api.Load_dataSQLiteAsyncLoaddata(sql)
                Dim dcResultdata As Object = New JavaScriptSerializer().Deserialize(Of List(Of Object))(jsonData)
                Dim tmp_id As String = ""
                For Each items As Object In dcResultdata
                    tmp_id = items("tmp_id").ToString()
                Next
                Return tmp_id
            Catch ex As Exception
                load_show.Show()
            End Try
        Catch ex As Exception
            MsgBox("Error Files model_api_sqlite In Function mas_Update_seqplan")
        End Try
    End Function
    Public Shared Async Function mas_manage_mold(line_cd As String) As Task(Of String)
        Dim date_start As String = DateTime.Now.ToString("yyyy/MM/dd") & " 08:00:00"
        Dim parsed_date_start As DateTime = DateTime.ParseExact(date_start, "yyyy/MM/dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
        Dim formatted_date_start As DateTime = parsed_date_start ' เก็บเป็น DateTime เพื่อใช้ AddDays ได้
        Dim convertStDate = formatted_date_start.ToString("yyyy-MM-dd HH:mm:ss")
        ' เพิ่ม 1 วัน
        Dim date_end As DateTime = formatted_date_start.AddDays(1)
        Dim convertdate_end = date_end.ToString("yyyy-MM-dd HH:mm:ss")
        Dim sqlGetact_ins = "SELECT * FROM act_ins WHERE line_cd = '" & line_cd & "' ORDER BY id DESC LIMIT 1;"
        Dim api = New api
        Dim jsonData As String = Await api.Load_dataSQLiteAsyncLoaddata(sqlGetact_ins)
        Try
            Dim dcResultdata As Object = New JavaScriptSerializer().Deserialize(Of List(Of Object))(jsonData)
            Dim tmp_seq_mold_no As String = ""
            Dim mm_id As String
            Dim imc_id As String
            Dim pwi_id As String
            Dim st_time As String
            Dim end_time As String
            Dim emp_code As String
            Dim ima_type_actual
            Dim ima_status_flg
            For Each items As Object In dcResultdata
                tmp_seq_mold_no = items("seq_mold_no").ToString()
                mm_id = items("mm_id").ToString
                imc_id = items("imc_id").ToString
                pwi_id = items("pwi_id").ToString
                st_time = items("st_time").ToString
                end_time = items("end_time").ToString
                emp_code = items("line_cd").ToString
                ima_type_actual = "2"
                ima_status_flg = "1"
            Next
            Dim sqlSum = "select IFNULL(SUM(qty), 0) as rs from act_ins where st_time >= '" & convertStDate & "' and end_time <= '" & convertdate_end & "' and line_cd ='" & line_cd & "' and seq_mold_no = '" & tmp_seq_mold_no & "'"
            'Console.WriteLine(sqlSum)
            Dim jsonDataSum As String = Await api.Load_dataSQLiteAsyncLoaddata(sqlSum)
            Dim dcResultdataSum As Object = New JavaScriptSerializer().Deserialize(Of List(Of Object))(jsonDataSum)
            Dim tmp_id As String = ""
            For Each items As Object In dcResultdataSum
                ' Dim Cavity = modelMoldCavity.GetCavity(mm_id)
                ' Dim ima_use_shot = Math.Ceiling(CDbl(Val(items("rs").ToString)) / Cavity)
                ' modelMoldCavity.mupdateShot(mm_id, pwi_id, ima_use_shot, ima_type_actual, st_time, end_time, ima_status_flg, emp_code, line_cd)
                ' modelMoldCavity.mUpdateStatusProduction("0", imc_id, line_cd, "1", "2")
            Next
        Catch ex As Exception
        End Try
        Return 0
    End Function
    Public Shared Async Function Get_time_start_shift(shift As String) As Task(Of DataTable)
        Try
            Dim api = New api()
            Dim sql As String = "EXEC [dbo].[GET_DATA_SHIFT] @shift = '" & shift & "'"
            Dim jsonData As String = Await api.Load_dataSQLiteAsyncLoaddata(sql)
            Dim result As DataTable = JsonConvert.DeserializeObject(Of DataTable)(jsonData)
            Return result
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Shared Async Function mas_Get_Plan_All_By_Line_Loss_E1(line_cd As String, shift As String, dateStart As String, timeStart As String, flg_spec As String, item_cd As String, TimeStartShift As String) As Task(Of String)
        Try
            Dim defaultTime As String = dateStart & " " & timeStart
            Dim dateTimeStartShift As String = ""
            Dim start_date As String = ""

            ' คำนวณเวลาเริ่มตามเงื่อนไข shift
            If String.Compare(timeStart, "00:00:00") >= 0 AndAlso String.Compare(timeStart, "07:59:59") <= 0 Then
                Dim delDate As String = DateTime.Parse(dateStart).AddDays(-1).ToString("yyyy-MM-dd")
                start_date = delDate & " " & TimeStartShift
            Else
                start_date = dateStart & " " & TimeStartShift
                If DateTime.Parse(start_date) > DateTime.Parse(defaultTime) Then
                    start_date = defaultTime
                End If
            End If
            dateTimeStartShift = start_date
            Dim dt_start As DateTime = DateTime.Parse(start_date)
            Dim dt_end As DateTime = DateTime.Parse(defaultTime)
            Dim diffDateTimeBefore11_min As Integer = CInt((dt_end - dt_start).TotalMinutes)
            ' หากมากกว่า 10 นาทีให้เพิ่มเวลาอีก 11 นาที
            If diffDateTimeBefore11_min > 10 Then
                start_date = dt_start.AddMinutes(11).ToString("yyyy-MM-dd HH:mm:ss")
            End If
            Dim api = New api()
            Dim sql As String
            Dim sql_check_loss As String
            ' สร้าง SQL ตาม flg_spec
            If flg_spec = "1" Then
                sql = "SELECT * FROM act_ins WHERE " &
                  "st_time BETWEEN '" & dateTimeStartShift & "' AND '" & defaultTime & "' AND " &
                  "line_cd = '" & line_cd & "' AND item_cd = '" & item_cd & "' AND qty > 0 " &
                  "ORDER BY end_time DESC LIMIT 1;"
                sql_check_loss = "SELECT * FROM loss_actual WHERE " &
                  "(line_cd = '" & line_cd & "' AND start_loss BETWEEN '" & dateTimeStartShift & "' AND '" & defaultTime & "' AND item_cd = '" & item_cd & "' AND loss_cd_id <> '1' AND flg_control <> '2') OR " &
                  "(line_cd = '" & line_cd & "' AND end_loss BETWEEN '" & dateTimeStartShift & "' AND '" & defaultTime & "' AND item_cd = '" & item_cd & "' AND loss_cd_id <> '1' AND flg_control <> '2') " &
                  "ORDER BY id DESC LIMIT 1;"
            Else
                sql = "SELECT * FROM act_ins WHERE " &
                  "st_time BETWEEN '" & dateTimeStartShift & "' AND '" & defaultTime & "' AND " &
                  "line_cd = '" & line_cd & "' AND qty > 0 " &
                  "ORDER BY end_time DESC LIMIT 1;"
                sql_check_loss = "SELECT * FROM loss_actual WHERE " &
                  "(line_cd = '" & line_cd & "' AND start_loss BETWEEN '" & dateTimeStartShift & "' AND '" & defaultTime & "' AND loss_cd_id <> '1' AND flg_control <> '2') OR " &
                  "(line_cd = '" & line_cd & "' AND end_loss BETWEEN '" & dateTimeStartShift & "' AND '" & defaultTime & "' AND loss_cd_id <> '1' AND flg_control <> '2') " &
                  "ORDER BY id DESC LIMIT 1;"
            End If
            ' โหลดข้อมูลการผลิต
            Dim jsonDataProd As String = Await api.Load_dataSQLiteAsyncLoaddata(sql)
            If String.IsNullOrWhiteSpace(jsonDataProd) OrElse Not jsonDataProd.Trim().StartsWith("[") Then
                Console.WriteLine("jsonDataProd not valid: " & jsonDataProd)
                jsonDataProd = "[]"
            End If
            Dim getProd As DataTable = Nothing
            Try
                getProd = JsonConvert.DeserializeObject(Of DataTable)(jsonDataProd)
            Catch ex As Exception
                MsgBox("Error parsing jsonDataProd: " & ex.Message)
                Return "0"
            End Try
            ' โหลดข้อมูล loss
            Dim jsonDataLoss As String = Await api.Load_dataSQLiteAsyncLoaddata(sql_check_loss)
            If String.IsNullOrWhiteSpace(jsonDataLoss) OrElse Not jsonDataLoss.Trim().StartsWith("[") Then
                Console.WriteLine("jsonDataLoss not valid: " & jsonDataLoss)
                jsonDataLoss = "[]"
            End If
            Dim getLoss As DataTable = Nothing
            Try
                getLoss = JsonConvert.DeserializeObject(Of DataTable)(jsonDataLoss)
            Catch ex As Exception
                MsgBox("Error parsing jsonDataLoss: " & ex.Message)
                Return "0"
            End Try
            ' ตรวจสอบว่าไม่มีทั้ง production และ loss
            If getProd Is Nothing OrElse getProd.Rows.Count = 0 Then
                If getLoss Is Nothing OrElse getLoss.Rows.Count = 0 Then
                    Dim diffMinutes As Integer = CInt((DateTime.Parse(defaultTime) - DateTime.Parse(start_date)).TotalMinutes)
                    If diffDateTimeBefore11_min > 10 Then
                        Dim data = New List(Of Dictionary(Of String, Object)) From {
                        New Dictionary(Of String, Object) From {
                            {"Time_From", "Current"},
                            {"Start_Loss", start_date},
                            {"End_Loss", defaultTime},
                            {"Loss_Time", diffMinutes},
                            {"Loss_Code", "37"}
                        }
                    }
                        Dim jsonString As String = JsonConvert.SerializeObject(data)
                        Return jsonString
                    End If
                End If
            End If
            Return "0"
        Catch ex As Exception
            Return "ERROR: " & ex.Message
        End Try
    End Function
End Class