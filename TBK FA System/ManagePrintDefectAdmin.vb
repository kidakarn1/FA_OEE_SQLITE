Imports System.Web.Script.Serialization

Public Class ManagePrintDefectAdmin

    ' Flag ใช้ป้องกันการทำงานของ SelectedIndexChanged ตอนกำลัง Load ComboBox
    Private isLoaded As Boolean = False

    ' กดปุ่ม Back
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Me.Close()
    End Sub

    ' ตอนโหลดฟอร์ม
    Private Sub ManagePrintDefectAdmin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        combodfType()
        loadDataDefect("1", Show_reprint_wi.hide_wi_select.Text, Scan_reprint.date_now_start, Scan_reprint.date_now_end)
        isLoaded = True ' เปิดให้ SelectedIndexChanged ทำงานหลังโหลดเสร็จ
    End Sub

    ' โหลดข้อมูลให้ ComboBox
    Public Sub combodfType()
        Try
            comboxitemtype.Items.Clear()
        Catch ex As Exception

        End Try
        ' Key = ชื่อ, Value = รหัส
        Dim myItems As New List(Of KeyValuePair(Of String, Integer)) From {
            New KeyValuePair(Of String, Integer)("NG", 1),
            New KeyValuePair(Of String, Integer)("NC", 2)
        }

        ' Binding
        comboxitemtype.DataSource = New BindingSource(myItems, Nothing)
        comboxitemtype.DisplayMember = "Key"     ' แสดงคำว่า NC / NG
        comboxitemtype.ValueMember = "Value"     ' ค่าเป็น 1 หรือ 2
        comboxitemtype.SelectedIndex = 0         ' เลือกรายการแรกไว้ก่อน
    End Sub

    Public Sub loadDataDefect(df_item_type As String, df_wi As String, DateStart As String, DateEnd As String)
        Dim rsData = modelDefect.LoadDataTagDefect(df_item_type, df_wi, DateStart, DateEnd)
        If rsData <> "0" Then
            ' เคลียร์ ListView ก่อน
            lvShowDataDefect.BeginUpdate()
            lvShowDataDefect.Items.Clear()
            lvShowDataDefect.Columns.Clear()
            lvShowDataDefect.View = View.Details
            ' สร้างคอลัมน์
            lvShowDataDefect.Columns.Add("NO", 60)
            lvShowDataDefect.Columns.Add("Part NO", 190)
            lvShowDataDefect.Columns.Add("Lot No", 100)
            lvShowDataDefect.Columns.Add("SEQ", 80)
            lvShowDataDefect.Columns.Add("QTY", 80)
            lvShowDataDefect.Columns.Add("BOX", 80)
            lvShowDataDefect.Columns.Add("Shift", 80)
            ' แปลง JSON
            Dim jsonData As List(Of Object) = New JavaScriptSerializer().Deserialize(Of List(Of Object))(rsData)
            Dim No As Integer = 0
            ' วนลูปใส่ข้อมูล
            For Each itemObj As Object In jsonData
                Dim item As Dictionary(Of String, Object) = CType(itemObj, Dictionary(Of String, Object))
                No += 1

                Dim row As New ListViewItem(No.ToString()) ' NO
                row.SubItems.Add(item("dti_item_cd").ToString())       ' Part NO
                row.SubItems.Add(item("dti_lot_no").ToString())        ' Lot No
                row.SubItems.Add(item("dti_seq_no").ToString())        ' SEQ
                row.SubItems.Add(item("dti_sum_qty").ToString())       ' QTY
                row.SubItems.Add(item("dti_box_no").ToString())        ' BOX

                If item.ContainsKey("dti_shift") Then
                    row.SubItems.Add(item("dti_shift").ToString())     ' Shift
                Else
                    row.SubItems.Add("-")
                End If

                lvShowDataDefect.Items.Add(row)
            Next

            lvShowDataDefect.EndUpdate()

        Else
            ' ✅ กรณีไม่มีข้อมูล: ลบเฉพาะแถว แต่คงคอลัมน์ไว้
            lvShowDataDefect.BeginUpdate()
            lvShowDataDefect.Items.Clear()
            lvShowDataDefect.EndUpdate()
            MsgBox("No data found.", MsgBoxStyle.Information)
        End If

#If DEBUG Then
        MsgBox(rsData)
#End If

    End Sub




    ' เมื่อเลือกประเภท defect
    Private Sub comboxitemtype_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comboxitemtype.SelectedIndexChanged
        If Not isLoaded Then Exit Sub ' ยังไม่โหลดเสร็จ ไม่ต้องทำงาน
        ' ป้องกัน error ถ้า SelectedValue ยังไม่ถูกต้อง
        If comboxitemtype.SelectedValue IsNot Nothing AndAlso TypeOf comboxitemtype.SelectedValue Is Integer Then
            Dim selectedValue As Integer = Convert.ToInt32(comboxitemtype.SelectedValue)
            Dim selectedText As String = comboxitemtype.Text
            MessageBox.Show("คุณเลือก: " & selectedText & " (Key = " & selectedValue & ")")
            ' เรียกโหลดข้อมูล defect
            loadDataDefect(selectedValue, Show_reprint_wi.hide_wi_select.Text, Scan_reprint.date_now_start, Scan_reprint.date_now_end)
        End If
    End Sub
    ' ปุ่มอื่น (ว่าง)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' กำหนดโค้ดตามต้องการ
        PrintDefectAdmin()
    End Sub
    Public Sub PrintDefectAdmin()
        '  Dim objTagprintdefect = New printDefect()
        '  objTagprintdefect.Set_parameter_print(itemdf("dt_item_cd").ToString(), detailItemfg("ITEM_NAME").ToString(), detailItemfg("MODEL").ToString(), sLine, stDatetime, detailItemfg("LOCATION_PART").ToString(), sShift, factory_cd, sLot, itemdf("total_nc"), SEQ, wi, itemType, dfType, Menu)

    End Sub
    Private Sub lvShowDataDefect_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvShowDataDefect.SelectedIndexChanged
        If lvShowDataDefect.SelectedItems.Count > 0 Then
            Dim selectedItem As ListViewItem = lvShowDataDefect.SelectedItems(0)
            ' ดึงข้อมูลจากคอลัมน์ย่อย
            Dim partNo As String = selectedItem.SubItems(1).Text
            Dim lotNo As String = selectedItem.SubItems(2).Text
            Dim seq As String = selectedItem.SubItems(3).Text

            ' แสดงผล (หรือส่งไปยัง textbox, ตัวแปร, function พิมพ์ ฯลฯ)
            MsgBox("คุณเลือก Part No: " & partNo & vbCrLf &
                   "Lot No: " & lotNo & vbCrLf &
                   "SEQ: " & seq)
        End If
    End Sub


End Class
