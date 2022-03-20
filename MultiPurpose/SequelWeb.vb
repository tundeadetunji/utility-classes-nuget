Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Web.UI.WebControls
Imports MultiPurpose.UtilityWeb
Imports MultiPurpose.Utility

Public Class SequelWeb
	Public Sub DropTextBoolean(d_ As DropDownList, Optional pattern_ As String = "always/never", Optional FirstIsEmpty As Boolean = True)
		If d_.Items.Count > 0 Then Exit Sub

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		With d_.Items
			Select Case pattern_.Trim.ToLower
				Case ""
					.Add("Always")
					.Add("Never")
				Case "yes/no"
					.Add("Yes")
					.Add("No")
				Case "always/never"
					.Add("Always")
					.Add("Never")
				Case "on/off"
					.Add("On")
					.Add("Off")
				Case "1/0"
					.Add("1")
					.Add("0")
				Case "true/false"
					.Add("True")
					.Add("False")
			End Select

		End With

	End Sub

	Public Sub CData(c_ As CheckBoxList, query As String, connection_string As String, data_text_field As String, Optional data_value_field As String = "", Optional select_parameter_keys_values_ As Array = Nothing)
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_
		Try
			c_.DataSource = Nothing
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

			c_.DataSource = dt
			c_.DataTextField = data_text_field
			If data_value_field.Length > 0 Then c_.DataValueField = data_value_field
			c_.DataBind()
		Catch
		End Try


		'		d.GData(gPayment, Payment_, g_con)

		'		Dim select_parameter_keys_values() = {"AccountID", Context.User.Identity.GetUserName()}
		'		d.GData(gPayment, School_, m_con, select_parameter_keys_values)

	End Sub
	Public Sub ClearList(l_ As ListBox)
		Try
			l_.DataSource = Nothing
		Catch ex As Exception
		End Try
		l_.Items.Clear()
	End Sub

	Public Sub ClearDropDown(d_ As DropDownList)
		Try
			d_.DataSource = Nothing
		Catch ex As Exception
		End Try
		d_.Items.Clear()
	End Sub

	''' <summary>
	''' Binds SQL Server database table to GridView. Returns the GridView. Same as GData.
	''' </summary>
	''' <param name="g_">GridView to bind to.</param>
	''' <param name="query">The SQL query.</param>
	''' <param name="connection_string">The server connection string.</param>
	''' <param name="select_parameter_keys_values_">The SQL select parameters.</param>
	''' <return>g_</return>
	Public Shared Function Display(g_ As GridView, query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As GridView
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
			g_.DataBind()
		Catch
		End Try

		Return g_
		'		d.GData(gPayment, Payment_, g_con)

		'		Dim select_parameter_keys_values() = {"AccountID", Context.User.Identity.GetUserName()}
		'		d.GData(gPayment, School_, m_con, select_parameter_keys_values)

	End Function

	''' <summary>
	''' Binds DropDownList to SQL database column.
	''' </summary>
	''' <param name="d_"></param>
	''' <param name="query"></param>
	''' <param name="connection_string"></param>
	''' <param name="data_text_field"></param>
	''' <param name="data_value_field"></param>
	''' <param name="select_parameter_keys_values_"></param>
	''' <param name="DontIfFull">Ignores the function if d_ already has items</param>
	''' <returns>d_</returns>
	Public Shared Function DData(d_ As DropDownList, query As String, connection_string As String, data_text_field As String, Optional data_value_field As String = "", Optional select_parameter_keys_values_ As Array = Nothing, Optional DontIfFull As Boolean = False) As DropDownList
		If DontIfFull = True Then
			If d_.Items.Count > 0 Then Return d_
		End If
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
		d_.DataTextField = data_text_field
		If data_value_field.Length > 0 Then d_.DataValueField = data_value_field
		d_.DataBind()
		Return d_
	End Function

	''' <summary>
	''' Binds DropDownList to List.
	''' </summary>
	''' <param name="d_"></param>
	''' <param name="dont_if_full">Ignores the function if d_ already has items.</param>
	''' <param name="object_">List.</param>
	''' <returns>d_</returns>
	Public Shared Function DData(d_ As DropDownList, object_ As Object, Optional dont_if_full As Boolean = False) As DropDownList
		If dont_if_full And d_.Items.Count > 0 Then Return d_
		Try
			Dim l As New List(Of String)
			l = CType(object_, List(Of String))
			With d_
				.DataSource = l
				.DataBind()
			End With
		Catch
		End Try
		Return d_
	End Function

	Public Shared Function LData(l_ As ListBox, query As String, connection_string As String, data_text_field As String, Optional data_value_field As String = "", Optional select_parameter_keys_values_ As Array = Nothing) As ListBox
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
		l_.DataTextField = data_text_field
		If data_value_field.Length > 0 Then l_.DataValueField = data_value_field
		l_.DataBind()

		Return l_
	End Function


	''' <summary>
	''' Commits record to SQL Server database. Same as CommitRecord.
	''' </summary>
	''' <param name="query">The SQL query.</param>
	''' <param name="connection_string">The server connection string.</param>
	''' <param name="parameters_keys_values_">Values to put in table.</param>
	''' <returns>True if successful, False if not.</returns>
	Public Shared Function CommitSequel(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing) As Boolean
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = parameters_keys_values_
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
			Return False
		End Try

		'		Dim Entries_Insert As String = "INSERT INTO ENTRIES (EntryBy, ID, Category, [Description], Flag, [Title], Entry, DateAdded, TimeAdded, TitleID, Picture, PictureExtension, Topic) VALUES (@EntryBy, @ID, @Category, [@Description], @Flag, [@Title], @Entry, @DateAdded, @TimeAdded, @TitleID, @Picture, @PictureExtension, @Topic)"
		'		Dim entries_parameters_() = {"EntryBy", TitleBy.Text.Trim, "ID", EntryID.Text.Trim, "Category", Category.Text.Trim, "[Description]", Description.Text.Trim, "Flag", cFlag.Text.Trim, "[Title]", EntryTitle.Text.Trim, "Entry", NewEntry.Text.Trim, "DateAdded", date_, "TimeAdded", time_, "TitleID", TitleID.Text.Trim, "Picture", stream.GetBuffer(), "PictureExtension", PictureExtension.Text.Trim, "Topic", Topic.Text.Trim}
		'		d.CommitRecord(Entries_Insert, m_con, entries_parameters_)

	End Function

	Public Shared Function PictureFromStream(picture_ As System.Drawing.Image, Optional file_extension As String = ".jpg") As Byte()
		Dim photo_ As System.Drawing.Image = picture_
		Dim stream_ As New IO.MemoryStream

		If photo_ IsNot Nothing Then
			Select Case LCase(file_extension)
				Case ".jpg"
					photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
				Case ".jpeg"
					photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
				Case ".png"
					photo_.Save(stream_, Imaging.ImageFormat.Png)
				Case ".gif"
					photo_.Save(stream_, Imaging.ImageFormat.Gif)
				Case ".bmp"
					photo_.Save(stream_, Imaging.ImageFormat.Bmp)
				Case ".tif"
					photo_.Save(stream_, Imaging.ImageFormat.Tiff)
				Case ".ico"
					photo_.Save(stream_, Imaging.ImageFormat.Icon)
				Case Else
					photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
			End Select
		End If
		Return stream_.GetBuffer
	End Function
#Region "Bindings"

	''' <summary>
	''' Attaches List(Of String) as DataSource to DropDownList 
	''' </summary>
	''' <param name="control_"></param>
	''' <param name="list_"></param>
	''' <param name="First_Item_Is_Empty"></param>
	''' <returns></returns>
	Public Shared Function BindProperty(control_ As DropDownList, list_ As List(Of String), Optional First_Items_Are As Array = Nothing, Optional First_Item_Is_Empty As Boolean = False, Optional clear_before_fill As Boolean = False) As WebControl
		If control_.Items.Count > 0 Then Return control_

		Dim l As New List(Of String)
		If First_Item_Is_Empty Then l.Add("")
		If First_Items_Are IsNot Nothing Then
			If First_Items_Are.Length > 0 Then
				With First_Items_Are
					For i = 0 To .Length - 1
						l.Add(First_Items_Are(i))
					Next
				End With
			End If
		End If
		For i = 0 To CType(list_, List(Of String)).Count - 1
			l.Add(list_(i))
		Next
		Try
			If clear_before_fill Then
				CType(control_, DropDownList).Items.Clear()
				CType(control_, DropDownList).DataSource = Nothing
			End If
		Catch ex As Exception

		End Try
		With l
			Try
				For i = 0 To .Count - 1
					CType(control_, DropDownList).Items.Add(l(i))
				Next
			Catch ex As Exception

			End Try
		End With
		Return control_
	End Function

	''' <summary>
	''' Attaches Array as DataSource to DropDownList 
	''' </summary>
	''' <param name="control_"></param>
	''' <param name="list_"></param>
	''' <param name="First_Item_Is_Empty"></param>
	''' <returns></returns>
	Public Shared Function BindProperty(control_ As DropDownList, list_ As Array, Optional First_Items_Are As Array = Nothing, Optional First_Item_Is_Empty As Boolean = False, Optional clear_before_fill As Boolean = False) As WebControl
		If control_.Items.Count > 0 Then Return control_

		Dim l As New List(Of String)
		If First_Item_Is_Empty Then l.Add("")
		If First_Items_Are IsNot Nothing Then
			If First_Items_Are.Length > 0 Then
				With First_Items_Are
					For i = 0 To .Length - 1
						l.Add(First_Items_Are(i))
					Next
				End With
			End If
		End If
		If TypeOf (list_) Is Array Then
			For i = 0 To CType(list_, Array).Length - 1
				l.Add(list_(i))
			Next
		End If

		Try
			If clear_before_fill Then
				CType(control_, DropDownList).Items.Clear()
				CType(control_, DropDownList).DataSource = Nothing
			End If
		Catch ex As Exception

		End Try
		With l
			Try
				For i = 0 To .Count - 1
					CType(control_, DropDownList).Items.Add(l(i))
				Next
			Catch ex As Exception

			End Try
		End With
		Return control_
	End Function



	''' <summary>
	''' Attaches List(Of String) as DataSource to ListBox
	''' </summary>
	''' <param name="control_"></param>
	''' <param name="list_"></param>
	''' <param name="First_Item_Is_Empty"></param>
	''' <returns></returns>
	Public Shared Function BindProperty(control_ As ListBox, list_ As List(Of String), Optional First_Items_Are As Array = Nothing, Optional First_Item_Is_Empty As Boolean = False, Optional clear_before_fill As Boolean = False) As WebControl
		If control_.Items.Count > 0 Then Return control_
		Dim l As New List(Of String)
		If First_Item_Is_Empty Then l.Add("")
		If First_Items_Are IsNot Nothing Then
			If First_Items_Are.Length > 0 Then
				With First_Items_Are
					For i = 0 To .Length - 1
						l.Add(First_Items_Are(i))
					Next
				End With
			End If
		End If
		For i = 0 To CType(list_, List(Of String)).Count - 1
			l.Add(list_(i))
		Next
		Try
			If clear_before_fill Then
				CType(control_, ListBox).Items.Clear()
				CType(control_, ListBox).DataSource = Nothing
			End If
		Catch ex As Exception
		End Try
		With l
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
	''' Attaches Array as DataSource to ListBox
	''' </summary>
	''' <param name="control_"></param>
	''' <param name="list_"></param>
	''' <param name="First_Item_Is_Empty"></param>
	''' <returns></returns>
	Public Shared Function BindProperty(control_ As ListBox, list_ As Array, Optional First_Items_Are As Array = Nothing, Optional First_Item_Is_Empty As Boolean = False, Optional clear_before_fill As Boolean = False) As WebControl
		If control_.Items.Count > 0 Then Return control_
		Dim l As New List(Of String)
		If First_Item_Is_Empty Then l.Add("")
		If First_Items_Are IsNot Nothing Then
			If First_Items_Are.Length > 0 Then
				With First_Items_Are
					For i = 0 To .Length - 1
						l.Add(First_Items_Are(i))
					Next
				End With
			End If
		End If
		If TypeOf (list_) Is Array Then
			For i = 0 To CType(list_, Array).Length - 1
				l.Add(list_(i))
			Next
		End If
		Try
			If clear_before_fill Then
				CType(control_, ListBox).Items.Clear()
				CType(control_, ListBox).DataSource = Nothing
			End If
		Catch ex As Exception
		End Try
		With l
			Try
				For i = 0 To .Count - 1
					CType(control_, ListBox).Items.Add(l(i))
				Next
			Catch ex As Exception

			End Try
		End With
		Return control_
	End Function



	'''' <summary>
	'''' Attaches List(Of String) as DataSource to DropDownList
	'''' </summary>
	'''' <param name="control_"></param>
	'''' <param name="list_"></param>
	'''' <param name="First_Item_Is_Empty"></param>
	'''' <returns></returns>
	'Public Shared Function BindProperty(control_ As DropDownList, list_ As List(Of String), Optional First_Item_Is As String = "", Optional First_Item_Is_Empty As Boolean = True) As WebControl
	'	Dim l As New List(Of String)
	'	If First_Item_Is_Empty Then l.Add("")
	'	If First_Item_Is.Trim.Length > 1 Then l.Add(First_Item_Is)
	'	For i = 0 To CType(list_, List(Of String)).Count - 1
	'		l.Add(list_(i))
	'	Next
	'	Try
	'		CType(control_, DropDownList).Items.Clear()
	'		CType(control_, DropDownList).DataSource = Nothing
	'	Catch ex As Exception

	'	End Try
	'	With l
	'		Try
	'			For i = 0 To .Count - 1
	'				CType(control_, DropDownList).Items.Add(l(i))
	'			Next
	'		Catch ex As Exception

	'		End Try
	'	End With
	'	Return control_
	'End Function

	'''' <summary>
	'''' Attaches Array as DataSource to DropDownList
	'''' </summary>
	'''' <param name="control_"></param>
	'''' <param name="list_"></param>
	'''' <param name="First_Item_Is_Empty"></param>
	'''' <returns></returns>
	'Public Shared Function BindProperty(control_ As DropDownList, list_ As Array, Optional First_Item_Is As String = "", Optional First_Item_Is_Empty As Boolean = True) As WebControl
	'	Dim l As New List(Of String)
	'	If First_Item_Is_Empty Then l.Add("")
	'	If First_Item_Is.Trim.Length > 1 Then l.Add(First_Item_Is)
	'	If TypeOf (list_) Is Array Then
	'		For i = 0 To CType(list_, Array).Length - 1
	'			l.Add(list_(i))
	'		Next
	'	End If
	'	Try
	'		CType(control_, DropDownList).Items.Clear()
	'		CType(control_, DropDownList).DataSource = Nothing
	'	Catch ex As Exception

	'	End Try
	'	With l
	'		Try
	'			For i = 0 To .Count - 1
	'				CType(control_, DropDownList).Items.Add(l(i))
	'			Next
	'		Catch ex As Exception

	'		End Try
	'	End With
	'	Return control_
	'End Function



	'''' <summary>
	'''' Attaches List(Of String) as DataSource to ListBox
	'''' </summary>
	'''' <param name="control_"></param>
	'''' <param name="list_"></param>
	'''' <param name="First_Item_Is_Empty"></param>
	'''' <returns></returns>
	'Public Shared Function BindProperty(control_ As ListBox, list_ As List(Of String), Optional First_Item_Is As String = "", Optional First_Item_Is_Empty As Boolean = True) As WebControl
	'	Dim l As New List(Of String)
	'	If First_Item_Is_Empty Then l.Add("")
	'	If First_Item_Is.Trim.Length > 1 Then l.Add(First_Item_Is)
	'	For i = 0 To CType(list_, List(Of String)).Count - 1
	'		l.Add(list_(i))
	'	Next
	'	Try
	'		CType(control_, ListBox).Items.Clear()
	'		CType(control_, ListBox).DataSource = Nothing
	'	Catch ex As Exception
	'	End Try
	'	With l
	'		Try
	'			For i = 0 To .Count - 1
	'				CType(control_, ListBox).Items.Add(l(i))
	'			Next
	'		Catch ex As Exception

	'		End Try
	'	End With
	'	Return control_
	'End Function

	'''' <summary>
	'''' Attaches Array as DataSource to ListBox
	'''' </summary>
	'''' <param name="control_"></param>
	'''' <param name="list_"></param>
	'''' <param name="First_Item_Is_Empty"></param>
	'''' <returns></returns>
	'Public Shared Function BindProperty(control_ As ListBox, list_ As Array, Optional First_Item_Is As String = "", Optional First_Item_Is_Empty As Boolean = True) As WebControl
	'	Dim l As New List(Of String)
	'	If First_Item_Is_Empty Then l.Add("")
	'	If First_Item_Is.Trim.Length > 1 Then l.Add(First_Item_Is)
	'	If TypeOf (list_) Is Array Then
	'		For i = 0 To CType(list_, Array).Length - 1
	'			l.Add(list_(i))
	'		Next
	'	End If
	'	Try
	'		CType(control_, ListBox).Items.Clear()
	'		CType(control_, ListBox).DataSource = Nothing
	'	Catch ex As Exception
	'	End Try
	'	With l
	'		Try
	'			For i = 0 To .Count - 1
	'				CType(control_, ListBox).Items.Add(l(i))
	'			Next
	'		Catch ex As Exception

	'		End Try
	'	End With
	'	Return control_
	'End Function


	''' <summary>
	''' Binds database column/field to DropDownList.
	''' </summary>
	''' <param name="control_">Control</param>
	''' <param name="query_">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="select_parameter_keys_values_">Select Parameters (Nothing, if not needed)</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="First_Item_Is_Empty">Should first element of ComboBox appear empty?</param>
	''' <returns>control_</returns>
	Public Shared Function BindProperty(control_ As DropDownList, query_ As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing, Optional data_text_field As String = "", Optional First_Item_Is_Empty As Boolean = True, Optional clear_before_fill As Boolean = False) As WebControl
		If control_.Items.Count > 0 Then Return control_

		'c, text
		'c, checked

		'g
		'If TypeOf control_ Is GridView Then
		'	Return Display(control_, query_, connection_string, select_parameter_keys_values_)
		'End If
		'd, text
		'd, data
		If TypeOf control_ Is DropDownList Then
			Return DData(control_, query_, connection_string, data_text_field, data_text_field, select_parameter_keys_values_, First_Item_Is_Empty)
		End If
		'l
		'If TypeOf control_ Is ListBox Then
		'	Return LData(control_, query_, connection_string, data_text_field, data_text_field, select_parameter_keys_values_)
		'End If
	End Function

	''' <summary>
	''' Binds database column/field to ListBox.
	''' </summary>
	''' <param name="control_">Control</param>
	''' <param name="query_">SQL Query</param>
	''' <param name="connection_string">SQL Connection String</param>
	''' <param name="select_parameter_keys_values_">Select Parameters (Nothing, if not needed)</param>
	''' <param name="data_text_field">Database Field</param>
	''' <param name="First_Item_Is_Empty">Should first element of ComboBox appear empty?</param>
	''' <returns>control_</returns>
	Public Shared Function BindProperty(control_ As ListBox, query_ As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing, Optional data_text_field As String = "", Optional First_Item_Is_Empty As Boolean = True, Optional clear_before_fill As Boolean = False) As WebControl
		If control_.Items.Count > 0 Then Return control_
		'c, text
		'c, checked

		'g
		'If TypeOf control_ Is GridView Then
		'	Return Display(control_, query_, connection_string, select_parameter_keys_values_)
		'End If
		'd, text
		'd, data
		'If TypeOf control_ Is DropDownList Then
		'	Return DData(control_, query_, connection_string, data_text_field, data_text_field, select_parameter_keys_values_, First_Item_Is_Empty)
		'End If
		'l
		If TypeOf control_ Is ListBox Then
			Return LData(control_, query_, connection_string, data_text_field, data_text_field, select_parameter_keys_values_)
		End If
	End Function

#End Region

	Public Shared Function RandomColor(l As List(Of Integer)) As String
		Dim r As Integer = RandomList(0, 256, l)
		Dim g As Integer = RandomList(0, 256, l)
		Dim b As Integer = RandomList(0, 256, l)
		Dim return_ As String = "rgb(" & r & ", " & g & ", " & b & ")"
		Return return_
	End Function

	Public Shared Function RandomColor(l As List(Of Integer), Optional alpha_ As Byte = Nothing) As String
		Dim a As Byte = alpha_
		If a < 0 Then a = 0
		If a > 1 Then a = 1

		Dim r As Integer = RandomList(0, 256, l)
		Dim g As Integer = RandomList(0, 256, l)
		Dim b As Integer = RandomList(0, 256, l)
		Dim return_ As String = "rgba(" & r & ", " & g & ", " & b & ", " & a & ")"
		Return return_
	End Function




End Class
