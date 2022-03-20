Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class SequelDesktop

    Private Shared Function CommitSQLRecord(query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As Boolean
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            Dim insert_query As String = query
            Using insert_conn As New SqlConnection(connection_string)
                Using insert_comm As New SqlCommand()
                    With insert_comm
                        .Connection = insert_conn
                        .CommandTimeout = 0
                        .CommandType = CommandType.Text
                        .CommandText = insert_query
                        If select_parameter_keys_values IsNot Nothing Then
                            For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                                .Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                            Next
                        End If
                    End With
                    Try
                        insert_conn.Open()
                        insert_comm.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End Using
            End Using
            Return True
        Catch ex As Exception
        End Try

        '		Dim Entries_Insert As String = "INSERT INTO ENTRIES (EntryBy, ID, Category, [Description], Flag, [Title], Entry, DateAdded, TimeAdded, TitleID, Picture, PictureExtension, Topic) VALUES (@EntryBy, @ID, @Category, [@Description], @Flag, [@Title], @Entry, @DateAdded, @TimeAdded, @TitleID, @Picture, @PictureExtension, @Topic)"
        '		Dim entries_parameters_() = {"EntryBy", TitleBy.Text.Trim, "ID", EntryID.Text.Trim, "Category", Category.Text.Trim, "[Description]", Description.Text.Trim, "Flag", cFlag.Text.Trim, "[Title]", EntryTitle.Text.Trim, "Entry", NewEntry.Text.Trim, "DateAdded", date_, "TimeAdded", time_, "TitleID", TitleID.Text.Trim, "Picture", stream.GetBuffer(), "PictureExtension", PictureExtension.Text.Trim, "Topic", Topic.Text.Trim}
        '		d.CommitRecord(Entries_Insert, a_con, entries_parameters_)

    End Function

    ''' <summary>
    ''' Commits record to SQL Server database by default, or to MS Access database if DB_Is_SQL_ is set to false. Same as CommitRecord.
    ''' </summary>
    ''' <param name="query">The SQL query.</param>
    ''' <param name="connection_string">The server connection string.</param>
    ''' <param name="parameters_keys_values_">Values to put in table.</param>
    ''' <returns>True if successful, False if not.</returns>
    Public Shared Function CommitSequel(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing, Optional DB_Is_SQL_ As Boolean = True) As Boolean
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = parameters_keys_values_

        If DB_Is_SQL_ = True Then
            CommitSQLRecord(query, connection_string, select_parameter_keys_values)
            Return True
            Exit Function
        End If

        Try
            Dim insert_query As String = query
            Using insert_conn As New OleDbConnection(connection_string)
                Using insert_comm As New OleDbCommand()
                    With insert_comm
                        .Connection = insert_conn
                        .CommandText = insert_query
                        If select_parameter_keys_values IsNot Nothing Then
                            For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                                .Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                            Next
                        End If
                    End With
                    Try
                        insert_conn.Open()
                        insert_comm.ExecuteNonQuery()
                    Catch ex As Exception
                    End Try
                End Using
            End Using
            Return True
        Catch ex As Exception
        End Try

    End Function

#Region "Bindings"
    ''' <summary>
    ''' Displays data on DataGridView.
    ''' </summary>
    ''' <param name="g_">DataGridView to bind to</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <example>Display(DataGridView, SQL_Query, Connection_String, Select_Parameters)</example>
    ''' <returns>g_</returns>

    Public Shared Function Display(g_ As DataGridView, query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As DataGridView
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            g_.DataSource = Nothing
        Catch ex As Exception
        End Try

        Try

            Dim connection As New SqlConnection(connection_string)
            Dim sql As String = query

            Dim Command = New SqlCommand(sql, connection)

            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If

            Dim da As New SqlDataAdapter(Command)
            Dim dt As New DataTable
            da.Fill(dt)

            g_.DataSource = dt
        Catch
        End Try

        Return g_


        '		d.GData(gPayment, Payment_, g_con)

        '		Dim select_parameter_keys_values() = {"AccountID", Context.User.Identity.GetUserName()}
        '		d.GData(gPayment, School_, m_con, select_parameter_keys_values)

    End Function

    ''' <summary>
    ''' Binds ComboBox to database column.
    ''' </summary>
    ''' <param name="d_">ComboBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Column</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <param name="First_Element_Is_Empty">Should first element of ComboBox appear empty?</param>
    ''' <returns></returns>
    Public Shared Function DData(d_ As ComboBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing, Optional First_Element_Is_Empty As Boolean = True) As ComboBox

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            d_.DataSource = Nothing
        Catch ex As Exception

        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        d_.DataSource = dt
        d_.DisplayMember = data_text_field

        If First_Element_Is_Empty Then d_.SelectedIndex = -1
        Return d_
    End Function

    ''' <summary>
    ''' Binds ComboBox Text property to database field.
    ''' </summary>
    ''' <param name="d_">ComboBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>d_</returns>
    Public Shared Function DText(d_ As ComboBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As ComboBox
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            d_.DataSource = Nothing
        Catch ex As Exception

        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Text", dt, data_text_field)
        d_.DataBindings.Add(b)

        Return d_
    End Function

    ''' <summary>
    ''' Binds CheckBox Checked property to database field.
    ''' </summary>
    ''' <param name="h_">CheckBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>h_</returns>
    Public Shared Function HData(h_ As CheckBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As CheckBox

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Checked", dt, data_text_field)
        h_.DataBindings.Add(b)


        Return h_
    End Function

    ''' <summary>
    ''' Binds CheckBox Text property to database field.
    ''' </summary>
    ''' <param name="h_">CheckBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>h_</returns>
    Public Shared Function HText(h_ As CheckBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As CheckBox

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Text", dt, data_text_field)
        h_.DataBindings.Add(b)

        Return h_

    End Function

    ''' <summary>
    ''' Binds ListBox to database column.
    ''' </summary>
    ''' <param name="l_">ListBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Column</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>l_</returns>
    Public Shared Function LData(l_ As ListBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As ListBox
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            l_.DataSource = Nothing
        Catch ex As Exception

        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        l_.DataSource = dt
        l_.DisplayMember = data_text_field

        Return l_
    End Function

    ''' <summary>
    ''' Binds PictureBox Image property to database field.
    ''' </summary>
    ''' <param name="p_">PictureBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>p_</returns>
    Public Shared Function PImage(p_ As PictureBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As PictureBox
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            p_.Image = Nothing
        Catch ex As Exception
        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Image", dt, data_text_field, True)
        p_.DataBindings.Add(b)

        Return p_
    End Function

    ''' <summary>
    ''' Binds PictureBox BackgroundImage property to database field.
    ''' </summary>
    ''' <param name="p_">PictureBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>p_</returns>
    Public Shared Function PBackgroundImage(p_ As PictureBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As PictureBox
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_
        Try
            p_.BackgroundImage = Nothing
        Catch ex As Exception
        End Try


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("BackgroundImage", dt, data_text_field, True)
        p_.DataBindings.Add(b)

        Return p_
    End Function

    ''' <summary>
    ''' Binds Button Text property to database field.
    ''' </summary>
    ''' <param name="b_">Button</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>b_</returns>
    Public Shared Function BData(b_ As Button, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As Button
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_


        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Text", dt, data_text_field)
        b_.DataBindings.Add(b)

        Return b_
    End Function

    ''' <summary>
    ''' Binds DateTimePicker Value property to database field.
    ''' </summary>
    ''' <param name="date_">DateTimePicker</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>date_</returns>
    Public Shared Function DATEData(date_ As DateTimePicker, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As DateTimePicker
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Value", dt, data_text_field)
        date_.DataBindings.Add(b)

        Return date_
    End Function

    ''' <summary>
    ''' Binds Label Text property to database field.
    ''' </summary>
    ''' <param name="label_">Label</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>label_</returns>
    Public Shared Function LABELData(label_ As Label, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As Label

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Text", dt, data_text_field)
        label_.DataBindings.Add(b)


        Return label_

    End Function

    ''' <summary>
    ''' Binds TextBox Text property to database field.
    ''' </summary>
    ''' <param name="t_">TextBox</param>
    ''' <param name="query">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters</param>
    ''' <returns>t_</returns>
    Public Shared Function TData(t_ As TextBox, query As String, connection_string As String, data_text_field As String, Optional select_parameter_keys_values_ As Array = Nothing) As TextBox

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Dim connection As New SqlConnection(connection_string)
        Dim sql As String = query

        Dim Command = New SqlCommand(sql, connection)

        Try
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If
        Catch
        End Try

        Dim da As New SqlDataAdapter(Command)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim b As New Binding("Text", dt, data_text_field)
        t_.DataBindings.Add(b)


        Return t_

    End Function

    Public Enum PropertyToBind
        Text
        Image
        BackgroundImage
        Items
        Check
    End Enum

    ''' <summary>
    ''' Attaches List(Of String) or Array as DataSource to ComboBox or ListBox
    ''' </summary>
    ''' <param name="control_"></param>
    ''' <param name="list_"></param>
    ''' <returns></returns>

    Public Shared Function BindProperty(control_ As Control, list_ As Object, Optional First_Item_Is As String = Nothing, Optional bindAsDatasourceNotList As Boolean = True) As Control
        If bindAsDatasourceNotList Then
            Try
                CType(control_, ComboBox).DataSource = list_
            Catch ex As Exception

            End Try

            Try
                CType(control_, ListBox).DataSource = list_
            Catch ex As Exception

            End Try
            Return control_
        End If

        Dim l As New List(Of String)
        'If First_Item_Is_Empty Then l.Add("")
        If First_Item_Is IsNot Nothing Then
            If First_Item_Is.Trim.Length > 1 Then
                l.Add(First_Item_Is)
            End If
        End If
        If TypeOf (list_) Is Array Then
            For i = 0 To CType(list_, Array).Length - 1
                l.Add(list_(i))
            Next
        ElseIf TypeOf (list_) Is List(Of String) Then
            For i = 0 To CType(list_, List(Of String)).Count - 1
                l.Add(list_(i))
            Next
        End If
        Try
            CType(control_, ComboBox).Items.Clear()
            CType(control_, ComboBox).DataSource = Nothing
        Catch ex As Exception

        End Try
        Try
            CType(control_, ListBox).Items.Clear()
            CType(control_, ListBox).DataSource = Nothing
        Catch ex As Exception
        End Try
        With l
            Try
                For i = 0 To .Count - 1
                    CType(control_, ComboBox).Items.Add(l(i))
                Next
            Catch ex As Exception

            End Try
            Try
                For i = 0 To .Count - 1
                    CType(control_, ListBox).Items.Add(l(i))
                Next
            Catch ex As Exception

            End Try
        End With
        Return control_
    End Function


    ''' <summary>
    ''' Binds Control property to database column/field.
    ''' </summary>
    ''' <param name="control_">Control</param>
    ''' <param name="property__">Property to bind to (Text, Value, Image, BackgroundImage or empty to bind to column)</param>
    ''' <param name="query_">SQL Query</param>
    ''' <param name="connection_string">SQL Connection String</param>
    ''' <param name="select_parameter_keys_values_">Select Parameters (Nothing, if not needed)</param>
    ''' <param name="data_text_field">Database Field</param>
    ''' <param name="First_Item_Is_Empty">Should first element of ComboBox appear empty?</param>
    ''' <returns>control_</returns>
    Public Shared Function BindProperty(control_ As Control, property__ As PropertyToBind, query_ As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing, Optional data_text_field As String = "", Optional First_Item_Is_Empty As Boolean = True) As Control
        Dim property_ As String = ""
        Select Case property__
            Case PropertyToBind.BackgroundImage
                property_ = property__.ToString.ToLower
            Case PropertyToBind.Check
                property_ = "checked"
            Case PropertyToBind.Image
                property_ = property__.ToString.ToLower
            Case PropertyToBind.Text
                property_ = "text"
        End Select

        'c, text
        'c, checked
        If TypeOf control_ Is CheckBox Then
            If property_.ToLower = "text" Then
                Return HText(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
            ElseIf property_.ToLower = "checked" Then
                Return HData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
            End If
        End If
        'g
        If TypeOf control_ Is DataGridView Then
            Return Display(control_, query_, connection_string, select_parameter_keys_values_)
        End If
        'd, text
        'd, data
        If TypeOf control_ Is ComboBox Then
            If property_.ToLower = "text" Then
                Return DText(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
            Else
                Return DData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_, First_Item_Is_Empty)
            End If
        End If
        'l
        If TypeOf control_ Is ListBox Then
            Return LData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
        End If
        'p, image
        'p, backgroundImage
        If TypeOf control_ Is PictureBox Then
            If property_.ToLower = "image" Then
                Return PImage(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
            Else
                Return PBackgroundImage(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
            End If
        End If
        'b, text
        If TypeOf control_ Is Button Then
            Return BData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
        End If
        'date, value
        If TypeOf control_ Is DateTimePicker Then
            Return DATEData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
        End If
        'l, text
        If TypeOf control_ Is Label Then
            Return LABELData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
        End If
        't, text
        If TypeOf control_ Is TextBox Then
            Return TData(control_, query_, connection_string, data_text_field, select_parameter_keys_values_)
        End If
    End Function

#End Region





End Class
