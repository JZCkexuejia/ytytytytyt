Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Class Form1
    Private connectionString As String = "Data Source=172.20.1.25;Connection Timeout=60000; Initial Catalog=MFG_ReportSystem_PRD;User ID=MFG_ReportSystem;Password=U2!5J9^45H35|N@YRCS2;pooling=true"
    Private Sub btnAddRow_Click(sender As Object, e As EventArgs) Handles btnAddRow.Click
        AddRow()
        LoadDataGridView()
    End Sub


    '加载'
    Private Sub LoadDataGridView()
        Dim query As String = "SELECT * FROM TestVB"
        lblRecord.Text = ""
        ' 创建一个新的SqlConnection对象
        Using connection As New SqlConnection(connectionString)
            Try
                ' 打开连接
                connection.Open()

                ' 创建一个SqlCommand对象
                Using command As New SqlCommand(query, connection)

                    ' 创建一个SqlDataAdapter对象来填充DataSet
                    Using adapter As New SqlDataAdapter(command)
                        Dim dataTable As New DataTable()
                        adapter.Fill(dataTable) ' 填充DataTable

                        ' 将DataTable设置为DataGridView的数据源
                        DataGridView1.DataSource = Nothing

                        DataGridView1.DataSource = dataTable
                        DataGridView1.AllowUserToAddRows = False

                    End Using ' SqlDataAdapter

                End Using ' SqlCommand

            Catch ex As Exception
                ' 异常处理
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using ' SqlConnection
    End Sub

    Private Sub AddRow()

        '判断是否存在'
        Dim Selects As String = "SELECT * FROM TestVB where Component = @FirstName"

        Dim insert As String = "INSERT INTO TestVB ([Component]) VALUES (@FirstName)"

        ' 创建一个新的SqlConnection对象
        Using connection As New SqlConnection(connectionString)
            Try
                ' 打开连接
                connection.Open()

                Dim selectCommand As New SqlCommand(Selects, connection)
                selectCommand.Parameters.AddWithValue("@FirstName", txtRow.Text)
                Using adapter As New SqlDataAdapter(selectCommand)
                    Dim dataTable As New DataTable()
                    adapter.Fill(dataTable) ' 填充DataTable
                    If dataTable.Rows.Count > 0 Then
                        MessageBox.Show("Error: 已存在")
                        Return
                    End If
                End Using ' adapter
                ' 创建一个SqlCommand对象
                Dim insertCommand As New SqlCommand(insert, connection)

                ' 添加参数  
                insertCommand.Parameters.AddWithValue("@FirstName", txtRow.Text)

                ' 执行命令  
                insertCommand.ExecuteNonQuery()

            Catch ex As Exception
                ' 异常处理
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using ' SqlConnection
    End Sub




    Private Sub AddColumn()
        Dim Max As String = txtMax.Text + "_Max"
        Dim Min As String = txtMin.Text + "_Min"

        Dim sb As New StringBuilder()
        sb.Append("IF not  EXISTS (  ")
        sb.Append("SELECT 1  FROM INFORMATION_SCHEMA.COLUMNS  WHERE TABLE_NAME = 'TestVB' AND COLUMN_NAME = '" + Max + "'  )  ")
        sb.Append("BEGIN  ALTER TABLE TestVB  ADD " + Max + " float null;ALTER TABLE TestVB  ADD " + Min + " float null; END ")
        Dim insert As String = sb.ToString()


        ' 创建一个新的SqlConnection对象
        Using connection As New SqlConnection(connectionString)
            Try
                ' 打开连接
                connection.Open()

                ' 创建一个SqlCommand对象
                Dim insertCommand As New SqlCommand(insert, connection)

                ' 执行命令  
                insertCommand.ExecuteNonQuery()

            Catch ex As Exception
                ' 异常处理
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using ' SqlConnection
    End Sub


    Private Sub btnAddColumn_Click(sender As Object, e As EventArgs) Handles btnAddColumn.Click
        AddColumn()
        LoadDataGridView()
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        LoadDataGridView()
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        LoadDataGridView()

        Dim number As Integer
        Dim contentnum As Double
        If Integer.TryParse(txtR.Text, number) = False Then
            ' 输入是有效的整数  
            MessageBox.Show("行必须输入的是数字：" & number.ToString())
            Return
        End If

        If Integer.TryParse(txtC.Text, number) = False Then
            ' 输入是有效的整数  
            MessageBox.Show("列必须输入的是数字：" & number.ToString())
            Return
        End If


        Dim pattern As String = "^\d+(\.\d+)?$" ' 匹配整数或小数（整数部分至少一位，小数部分可选） 

        If Regex.IsMatch(txtContent.Text, pattern) = False Then
            ' 输入是有效的数字或小数  
            MessageBox.Show("内容请输入数字或者小数")
            Return
        End If

        Dim rowIndex As Integer = txtR.Text
        Dim columnIndex As Integer = txtC.Text
        Dim content As Double = txtContent.Text
        Dim Msg As String = ""

        If rowIndex > DataGridView1.Rows.Count Then
            MessageBox.Show("行不存在")
            Return
        End If
        If columnIndex > DataGridView1.ColumnCount Then
            MessageBox.Show("列不存在")
            Return
        End If

        DataGridView1.Rows(rowIndex - 1).Cells(columnIndex - 1).Style.BackColor = Color.Yellow


        Dim result As DialogResult = MessageBox.Show("你确定要执行这个操作吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' 根据用户的响应执行操作  
        If result = DialogResult.No Then
            Return
        End If
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                Dim cmd As New SqlCommand("UpdateTestVB", conn)
                cmd.CommandType = CommandType.StoredProcedure

                ' 添加输入参数  
                cmd.Parameters.Add("@RowNum", SqlDbType.Int).Value = rowIndex
                cmd.Parameters.Add("@Col", SqlDbType.Int).Value = columnIndex
                cmd.Parameters.Add("@Content", SqlDbType.Float).Value = content


                Dim outputParam As New SqlParameter()
                outputParam.ParameterName = "@Msg"
                outputParam.Direction = ParameterDirection.Output
                outputParam.SqlDbType = SqlDbType.VarChar ' 或者根据你的存储过程输出参数的实际类型设置  
                outputParam.Size = 255 ' 或者更大的尺寸，根据你的需要  
                cmd.Parameters.Add(outputParam)


                ' 执行存储过程  
                Dim count As Integer = cmd.ExecuteNonQuery()

                ' 读取输出参数的值  
                Msg = cmd.Parameters("@Msg").Value
                MessageBox.Show(Msg)
                LoadDataGridView()
                DataGridView1.Rows(rowIndex - 1).Cells(columnIndex - 1).Style.BackColor = Color.Yellow
            End Using
        Catch ex As Exception
            ' 异常处理
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub btnUpdate2_Click(sender As Object, e As EventArgs) Handles btnUpdate2.Click

        LoadDataGridView()

        Dim number As Integer
        Dim contentnum As Double
        Dim pattern As String = "^\d+(\.\d+)?$" ' 匹配整数或小数（整数部分至少一位，小数部分可选） 
        Dim sb As StringBuilder = New StringBuilder()


        If Regex.IsMatch(txtContent2.Text, pattern) = False Then
            ' 输入是有效的数字或小数  
            MessageBox.Show("内容请输入数字或者小数")
            Return
        End If
        Dim row As String = txtRowName.Text
        Dim column As String = txtColName.Text
        Dim column1 As String = txtColName1.Text
        Dim content As Double = txtContent2.Text

        If column = "" And column1 = "" Then
            MessageBox.Show("请输入列名")
            Return
        End If

        Dim Msg As String = ""
        Dim rowIndex As Integer = -1 '后续如果还是等于-1就相当于没有找到单元格'
        Dim colIndex As Integer = -1
        Dim colIndex1 As Integer = -1

        For Each datarow As DataGridViewRow In DataGridView1.Rows
            ' 检查"Name"列的值是否匹配  
            If datarow.Cells("Component").Value.ToString() = row Then
                rowIndex = datarow.Index
                Exit For
            End If
        Next
        colIndex = DataGridView1.Columns.IndexOf(DataGridView1.Columns(column))
        colIndex1 = DataGridView1.Columns.IndexOf(DataGridView1.Columns(column1))

        If rowIndex = -1 Then
            MessageBox.Show("行不存在")
            Return
        End If
        '如果列不等于空但是找不到就要提示'
        If Not String.IsNullOrWhiteSpace(column) And colIndex = -1 Then
            MessageBox.Show("1列不存在")
            Return
        End If

        '如果列1不等于空但是找不到就要提示'
        If Not String.IsNullOrWhiteSpace(column1) And colIndex1 = -1 Then
            MessageBox.Show("2列不存在")
            Return
        End If

        '两个列都不等于空'
        If colIndex <> -1 And colIndex1 <> -1 Then
            DataGridView1.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.FromArgb(11, 11, 11, 1)
            sb.AppendFormat("update TestVB set {0}={1} ,{2}={1} where Component='{3}' ", column, txtContent2.Text, column1, row)

        ElseIf colIndex <> -1 And colIndex1 = -1 Then
            DataGridView1.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.FromArgb(11, 11, 11, 1)
            sb.AppendFormat("update TestVB set {0}={1} where Component='{2}'  ", column, txtContent2.Text, row)
        ElseIf colIndex = -1 And colIndex1 <> -1 Then
            DataGridView1.Rows(rowIndex).Cells(colIndex1).Style.BackColor = Color.FromArgb(11, 11, 11, 1)
            sb.AppendFormat("update TestVB set {0}={1} where Component='{2}' ", column1, txtContent2.Text, row)
        End If



        Dim result As DialogResult = MessageBox.Show("你确定要执行这个操作吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' 根据用户的响应执行操作  
        If result = DialogResult.No Then
            Return
        End If
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                Dim sql As String = sb.ToString()
                Dim insertCommand As New SqlCommand(sql, conn)

                ' 执行命令  
                Dim count As Integer = insertCommand.ExecuteNonQuery()
                If count > 0 Then
                    MessageBox.Show("修改成功！")
                    LoadDataGridView()
                    If colIndex <> -1 Then
                        DataGridView1.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.Blue
                    End If

                    If colIndex1 <> -1 Then
                        DataGridView1.Rows(rowIndex).Cells(colIndex1).Style.BackColor = Color.Blue
                    End If

                End If


            End Using
        Catch ex As Exception
            ' 异常处理
            MessageBox.Show("Error: " & ex.Message)
        End Try

    End Sub



    '行编辑事件'
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        ' 在这里添加你的代码  
        ' 例如，获取编辑后的值  
        Dim editedValue As String = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()

        Dim pattern As String = "^\d+(\.\d+)?$" ' 匹配整数或小数（整数部分至少一位，小数部分可选） 

        If Regex.IsMatch(editedValue, pattern) = False Then
            ' 输入是有效的数字或小数  
            MessageBox.Show("内容请输入数字或者小数")
            Return
        End If



        '获取列名如果列名为Max则需要和Min比较'
        Dim colIndex As String
        colIndex = DataGridView1.Columns(e.ColumnIndex).Name

        '如果包含MAX'，则找Min的列比较
        If colIndex.Contains("Max") Then
            Dim ob As Object = DataGridView1.Columns(colIndex.Replace("Max", "Min"))
            If ob IsNot Nothing Then
                'Min的值'
                Dim colIndexCompareName As String = DataGridView1.Rows(e.RowIndex).Cells(colIndex.Replace("Max", "Min")).Value.ToString()

                Dim colIndexCompare As Integer = If(colIndexCompareName = "", 0, Convert.ToInt32(colIndexCompareName))

                If Convert.ToDouble(editedValue) < Convert.ToDouble(colIndexCompare) And colIndexCompareName <> "" Then
                    MessageBox.Show("输入Max的值必须要大于MIN的值")

                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.Red
                    Dim strNew As String = ""
                    Dim str As String = lblRecord.Text
                    Dim arr() As String
                    arr = Split(str, "|")
                    For Each xx In arr
                        '如果不等于就重新赋值给strNew'
                        If xx <> e.RowIndex.ToString() & "," & e.ColumnIndex.ToString() Then
                            If strNew = "" Then
                                strNew = xx
                            Else
                                strNew += "|" & xx
                            End If
                        End If
                    Next

                    lblRecord.Text = strNew

                    Return
                End If
            End If
        ElseIf colIndex.Contains("Min") Then
            Dim ob As Object = DataGridView1.Columns(colIndex.Replace("Min", "Max"))
            If ob IsNot Nothing Then
                '否则就是Min的必须小于Max，以下是Max'
                Dim colIndexCompareName As String = DataGridView1.Rows(e.RowIndex).Cells(colIndex.Replace("Min", "Max")).Value.ToString()

                Dim colIndexCompare As Integer = If(colIndexCompareName = "", 0, Convert.ToInt32(colIndexCompareName))
                If Convert.ToDouble(editedValue) > Convert.ToDouble(colIndexCompare) And colIndexCompareName <> "" Then
                    MessageBox.Show("输入Min的值必须要小于Max的值")

                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.Red
                    Dim strNew As String = ""
                    Dim str As String = lblRecord.Text
                    Dim arr() As String
                    arr = Split(str, "|")
                    For Each xx In arr
                        '如果不等于就重新赋值给strNew'
                        If xx <> e.RowIndex.ToString() & "," & e.ColumnIndex.ToString() Then
                            If strNew = "" Then
                                strNew = xx
                            Else
                                strNew += "|" & xx
                            End If
                        End If
                    Next

                    lblRecord.Text = strNew

                    Return

                End If
            Else
                MessageBox.Show("只允许修改包含MAX或者Min")
                Return
            End If
        End If

        DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.FromArgb(199, 21, 133)

        'Dim Row As String = DataGridView1.Rows(e.RowIndex).Cells("Component").Value.ToString()
        '把保存过的记录'
        If lblRecord.Text = "" Then
            'lblRecord.Text = Row & e.RowIndex.ToString() & "," & e.ColumnIndex.ToString()
            lblRecord.Text = e.RowIndex.ToString() & "," & e.ColumnIndex.ToString()
        Else
            'lblRecord.Text += "|" & Row & "," & e.RowIndex.ToString() & "," & e.ColumnIndex.ToString()
            lblRecord.Text += "|" & e.RowIndex.ToString() & "," & e.ColumnIndex.ToString()
        End If
    End Sub

    Private Sub btnUpdateAll_Click(sender As Object, e As EventArgs) Handles btnUpdateAll.Click

        If lblRecord.Text = "" Then
            MessageBox.Show("请输入你需要修改的列")
            Return
        End If
        Dim sb As StringBuilder = New StringBuilder()
        Dim str As String = lblRecord.Text
        Dim arr() As String
        arr = Split(str, "|")

        sb.Append("begin")

        ' 输出数组元素  
        For Each x In arr
            '获取行和列的数字，然后保存'
            Dim Row As String = DataGridView1.Rows(x.Split(",")(0)).Cells("Component").Value.ToString()
            Dim col As String = DataGridView1.Rows(x.Split(",")(0)).Cells(Convert.ToInt32(x.Split(",")(1))).Value.ToString()
            Dim colName As String = DataGridView1.Columns(Convert.ToInt32(x.Split(",")(1))).Name
            sb.AppendFormat(" update TestVB set {0}={1} where Component='{2}';  ", colName, col, Row)
        Next
        sb.Append("end")


        Dim result As DialogResult = MessageBox.Show("你确定要执行这个操作吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' 根据用户的响应执行操作  
        If result = DialogResult.No Then
            Return
        End If

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                Dim sql As String = sb.ToString()
                Dim insertCommand As New SqlCommand(sql, conn)

                ' 执行命令  
                Dim count As Integer = insertCommand.ExecuteNonQuery()
                If count > 0 Then
                    MessageBox.Show("修改成功！")
                    LoadDataGridView()
                End If


            End Using
        Catch ex As Exception
            ' 异常处理
            MessageBox.Show("Error: " & ex.Message)
        End Try

    End Sub
End Class
