Imports System.Data.SqlClient
Imports MultiPurpose.Utility
Public Class Sequel
#Region ""
	Public Enum CommitDataTo
		CSV = 0
		Sequel = 1
	End Enum

	Public Structure CommitDataTargetInfo
		Public commit_data_to As CommitDataTo
		Public filename As String
		'Public table As String
		'Public connection_string As String

	End Structure

#End Region

#Region "Query Strings"

	Public Enum Queries
		DeleteString_CONDITIONAL
		BuildSelectString_BETWEEN
		BuildSelectString_LIKE
		BuildUpdateString_CONDITIONAL
		BuildCountString_CONDITIONAL
		BuildCountString_GROUPED_CONDITIONAL
		BuildSelectString_CONDITIONAL
		BuildSelectString_DISTINCT
		BuildInsertString
		BuildUpdateString
		BuildSelectString
		BuildCountString
		BuildCountString_GROUPED
		BuildSumString_GROUPED
		BuildSumString_GROUPED_CONDITIONAL
		BuildAVGString_GROUPED
		BuildAVGString_GROUPED_CONDITIONAL
		BuildMinString_GROUPED
		BuildMinString_GROUPED_CONDITIONAL
		BuildTopString
		BuildMaxString_GROUPED
		BuildMaxString_GROUPED_CONDITIONAL
		BuildMaxString
		BuildAVGString_UNGROUPED
		BuildAVGString_UNGROUPED_CONDITIONAL
		BuildSumString_UNGROUPED
		BuildSumString_UNGROUPED_CONDITIONAL
		BuildCountString_UNGROUPED
		BuildCountString_UNGROUPED_CONDITIONAL
		BuildSelectString_DISTINCT_BETWEEN_CONDITIONAL
		BuildAVGString_GROUPED_BETWEEN
		BuildSumString_GROUPED_BETWEEN
		BuildCountString_GROUPED_BETWEEN
		BuildTopString_CONDITIONAL
	End Enum

	Public Enum OrderBy
		DESC
		ASC
	End Enum


	''' <summary>
	''' Builds SQL Delete Query String.
	''' </summary>
	''' <param name="t_">Table to delete from.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to delete or not, followed by the operator to apply on the key.</param>
	''' <returns>String</returns>
	Public Shared Function DeleteString_CONDITIONAL(t_ As String, where_key_operator As Array) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "DELETE "

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			Else
				v = ""
			End If
			v &= ")"
		Else
			v = ""
		End If
		Return v
	End Function

	''' <summary>
	''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="select_params">Columns to select.</param>
	''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO">Columns to consider with condition, appended automatically with _FROM or _TO (as must be manually appended to respective Where Keys when used with Display.</param>
	''' <param name="OrderByField">Column to order by.</param>
	''' <param name="order_by">Ascending or Descending</param>
	''' <returns></returns>
	Public Shared Function BuildSelectString_BETWEEN(t_ As String, Optional select_params As Array = Nothing, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC) As String

		Dim v As String = "SELECT "

		If select_params IsNot Nothing Then
			For i As Integer = 0 To select_params.Length - 1
				v &= select_params(i)
				If select_params.Length > 1 And i <> select_params.Length - 1 Then
					v &= ", "
				End If
			Next
		Else
			v &= " *"
		End If
		v &= " FROM " & t_

		If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
			If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
				v &= " WHERE "
				For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
					v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO)"
					If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
		End If
		If OrderByField IsNot Nothing Then
			v &= " ORDER BY " & OrderByField
		End If
		If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
		Return v
	End Function
	Public Enum LIKE_SELECT
		AND_
		OR_
	End Enum
	''' <summary>
	''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="select_params">Columns to select.</param>
	''' <param name="where_keys">Columns to apply condition on before selecting.</param>
	''' <param name="OrderByField">Column to order by.</param>
	''' <param name="order_by">Ascending or Descending</param>
	''' <returns></returns>
	Public Shared Function BuildSelectString_LIKE(t_ As String, Optional select_params As Array = Nothing, Optional where_keys As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional like_operator As LIKE_SELECT = LIKE_SELECT.AND_) As String
		Dim v As String = "SELECT "

		If select_params IsNot Nothing Then
			For i As Integer = 0 To select_params.Length - 1
				v &= select_params(i)
				If select_params.Length > 1 And i <> select_params.Length - 1 Then
					v &= ", "
				End If
			Next
		Else
			v &= " *"
		End If
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & " LIKE '%' + @" & where_keys(j) & " + '%'"
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " " & like_operator.ToString.Replace("_", "") & " "
					End If
				Next
			End If
			v &= ")"
		End If
		If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
		If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
		Return v
	End Function


	''' <summary>
	''' Builds SQL Update Query String.
	''' </summary>
	''' <param name="t_">Table to update.</param>
	''' <param name="update_keys">Fields to replace.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to update or not, followed by the operator to apply on the key.</param>
	''' <returns>String.</returns>
	Public Shared Function BuildUpdateString_CONDITIONAL(t_ As String, update_keys As Array, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "UPDATE " & t_ & " SET "

		For j As Integer = 0 To update_keys.Length - 1
			v &= update_keys(j) & "=@" & update_keys(j)
			If update_keys.Length > 1 And j <> update_keys.Length - 1 Then
				v &= ", "
			End If
		Next

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("

				For k As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(k) & " " & operator_(where_keys(k + 1)) & " @" & where_keys(k)
					If where_keys.Length > 1 And k <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
				v &= ")"
			End If
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL Count Query String. Useful for Line Chart (e.g. combined with Where clause)
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to count or not, followed by the operator to apply on the key.</param>
	''' <returns>String</returns>
	Public Shared Function BuildCountString_CONDITIONAL(t_ As String, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT COUNT(*) "
		v &= " FROM " & t_

		Dim l As New List(Of String)
		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					If l.Contains(where_keys(j)) Then
						v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j) & "_" & l.LastIndexOf(where_keys(j))
					Else
						v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					End If
					l.Add(where_keys(j))
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function


	''' <summary>
	''' Builds SQL Count query string, grouped - useful for Pie chart.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to count or not, followed by the operator to apply on the key.</param>
	''' <param name="field_to_count">Field to count and group.</param>
	''' <returns>String</returns>
	Public Shared Function BuildCountString_GROUPED_CONDITIONAL(t_ As String, field_to_count As String, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT " & field_to_count & ", COUNT(*) "
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v & " GROUP BY " & field_to_count
	End Function

	Private Shared Function operator_(operator__ As String) As String
		If operator__ = "" Then
			Return "="
		Else
			Return operator__
		End If
	End Function

	Private Shared Function BuildSelectString_COMPLEX(t_ As String, Optional select_params As Array = Nothing, Optional where_key_operator As Array = Nothing, Optional where_keys_FOR_LIKE As Array = Nothing, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing)
		Dim where_keys As Array = where_key_operator
		Dim where_keys_like As Array = where_keys_FOR_LIKE
		Dim v As String = "SELECT "

		If select_params IsNot Nothing Then
			For i As Integer = 0 To select_params.Length - 1
				v &= select_params(i)
				If select_params.Length > 1 And i <> select_params.Length - 1 Then
					v &= ", "
				End If
			Next
		Else
			v &= " *"
		End If
		v &= " FROM " & t_

		If where_keys IsNot Nothing Or where_keys_like IsNot Nothing Or where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
			v &= " WHERE ("

			If where_keys IsNot Nothing Then
				If where_keys.Length > 0 Then
					For j As Integer = 0 To where_keys.Length - 1 Step 2
						v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
						If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
							v &= " AND "
						End If
					Next
				End If
			End If

			If where_keys_like IsNot Nothing Then
				If where_keys_like.Length > 0 Then
					If where_keys Is Nothing Then
						v &= ""
					Else
						v &= " AND "
					End If
					For j As Integer = 0 To where_keys_like.Length - 1
						v &= where_keys_like(j) & " LIKE '%' + @" & where_keys_like(j) & " + '%'"
						If where_keys_like.Length > 1 And j <> where_keys_like.Length - 1 Then
							v &= " AND "
						End If
					Next
				End If
			End If

			If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
				If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
					If where_keys Is Nothing And where_keys_like Is Nothing Then
						v &= ""
					ElseIf where_keys IsNot Nothing Or where_keys_like IsNot Nothing Then
						v &= " AND "
					End If
					For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
						v &= "" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO"
						If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
							v &= " AND "
						End If
					Next
				End If
			End If

			v &= ")"
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="select_params">Fields to select.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to select or not, followed by the operator to apply on the key.</param>
	''' <param name="OrderByField">Column to order by.</param>
	''' <param name="order_by">Ascending or Descending</param>
	''' <returns>String.</returns>
	Public Shared Function BuildSelectString_CONDITIONAL(t_ As String, Optional select_params As Array = Nothing, Optional where_key_operator As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT "

		If select_params IsNot Nothing Then
			For i As Integer = 0 To select_params.Length - 1
				v &= select_params(i)
				If select_params.Length > 1 And i <> select_params.Length - 1 Then
					v &= ", "
				End If
			Next
		Else
			v &= " *"
		End If
		v &= " FROM " & t_

		Dim l As New List(Of String)
		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					If l.Contains(where_keys(j)) Then
						v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j) & "_" & l.LastIndexOf(where_keys(j))
					Else
						v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					End If
					l.Add(where_keys(j))
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If
		If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
		If OrderByField IsNot Nothing Then v &= " " & order_by.ToString

		Return v
	End Function

	''' <summary>
	''' Builds SQL Select Query String with DISTINCT. Suitable for Reader. To use count instead, use BuildCountString.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="select_params">Fields to select.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
	''' <returns>String.</returns>

	Public Shared Function BuildSelectString_DISTINCT(t_ As String, select_params As Array, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT DISTINCT "

		If select_params IsNot Nothing Then
			For i As Integer = 0 To select_params.Length - 1
				v &= select_params(i)
				If select_params.Length > 1 And i <> select_params.Length - 1 Then
					v &= ", "
				End If
			Next
		Else
			v &= " *"
		End If
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL Insert Query String.
	''' </summary>
	''' <param name="t_">Table to insert into.</param>
	''' <param name="insert_keys">Columns to insert into.</param>
	''' <returns>String.</returns>
	Public Shared Function BuildInsertString(t_ As String, insert_keys As Array) As String
		Dim v As String = "INSERT INTO " & t_ & " ("

		For i As Integer = 0 To insert_keys.Length - 1
			v &= insert_keys(i)
			If insert_keys.Length > 1 And i <> insert_keys.Length - 1 Then
				v &= ", "
			End If
		Next

		v &= ") VALUES ("

		For j As Integer = 0 To insert_keys.Length - 1
			v &= "@" & insert_keys(j)
			If insert_keys.Length > 1 And j <> insert_keys.Length - 1 Then
				v &= ", "
			End If
		Next

		v &= ")"
		Return v
	End Function

	''' <summary>
	''' Builds SQL Update Query String.
	''' </summary>
	''' <param name="t_">Table to update.</param>
	''' <param name="update_keys">Fields to replace.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to update or not.</param>
	''' <returns>String.</returns>
	Public Shared Function BuildUpdateString(t_ As String, update_keys As Array, Optional where_keys As Array = Nothing) As String

		Dim v As String = "UPDATE " & t_ & " SET "

		For j As Integer = 0 To update_keys.Length - 1
			v &= update_keys(j) & "=@" & update_keys(j)
			If update_keys.Length > 1 And j <> update_keys.Length - 1 Then
				v &= ", "
			End If
		Next

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("

				For k As Integer = 0 To where_keys.Length - 1
					v &= where_keys(k) & "=@" & where_keys(k)
					If where_keys.Length > 1 And k <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
				v &= ")"
			End If
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="select_params">Fields to select.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
	''' <param name="OrderByField">Column to order by.</param>
	''' <param name="order_by">Ascending or Descending</param>
	''' <returns>String.</returns>
	Public Shared Function BuildSelectString(t_ As String, Optional select_params As Array = Nothing, Optional where_keys As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC) As String
		Dim v As String = "SELECT "

		If select_params IsNot Nothing Then
			For i As Integer = 0 To select_params.Length - 1
				v &= select_params(i)
				If select_params.Length > 1 And i <> select_params.Length - 1 Then
					v &= ", "
				End If
			Next
		Else
			v &= " *"
		End If
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If
		If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
		If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
		Return v
	End Function

	''' <summary>
	''' Builds SQL Count Query String.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to count or not.</param>
	''' <returns>String</returns>
	Public Shared Function BuildCountString(t_ As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT COUNT(*) "
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL Count query string, grouped - useful for Pie chart.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to count or not.</param>
	''' <param name="field_to_count">Field to count and group.</param>
	''' <returns>String</returns>
	Public Shared Function BuildCountString_GROUPED(t_ As String, field_to_count As String, field_to_group As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT " & field_to_count & ", COUNT(*) "
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v & " GROUP BY " & field_to_group
	End Function

	''' <summary>
	''' Builds SQL Sum query string, grouped - useful for Pie chart.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_sum">Field to sum.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to sum or not.</param>
	''' <returns>String</returns>
	Public Shared Function BuildSumString_GROUPED(t_ As String, field_to_group As String, field_to_sum As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT " & field_to_group & ", SUM(" & field_to_sum & ") "
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v & " GROUP BY " & field_to_group
	End Function

	''' <summary>
	''' Builds SQL Sum query string, grouped - useful for Pie chart.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_sum">Field to sum.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to sum or not, followed by the operator to apply on the key.</param>
	''' <returns>String</returns>
	Public Shared Function BuildSumString_GROUPED_CONDITIONAL(t_ As String, field_to_group As String, field_to_sum As String, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT " & field_to_group & ", SUM(" & field_to_sum & ") "

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v & " GROUP BY " & field_to_group
	End Function

	''' <summary>
	''' Builds SQL AVG query string, grouped - useful for Pie chart.
	''' </summary>
	''' <param name="t_">Table for fields to find AVG.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to sort for AVG or not.</param>
	''' <returns>String</returns>
	Public Shared Function BuildAVGString_GROUPED(t_ As String, field_to_group As String, field_to_apply_avg_on As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT " & field_to_group & ", AVG(" & field_to_apply_avg_on & ") "

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v & " GROUP BY " & field_to_group
	End Function

	''' <summary>
	''' Builds SQL AVG query string, grouped - useful for Pie chart.
	''' </summary>
	''' <param name="t_">Table for fields to find AVG.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to sort for AVG or not, followed by the operator to apply on the key.</param>
	''' <returns>String</returns>
	Public Shared Function BuildAVGString_GROUPED_CONDITIONAL(t_ As String, field_to_group As String, field_to_apply_avg_on As String, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT " & field_to_group & ", AVG(" & field_to_apply_avg_on & ") "

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v & " GROUP BY " & field_to_group
	End Function

	''' <summary>
	''' Builds SQL MIN query string, grouped - useful for Pie chart.
	''' </summary>
	''' <param name="t_">Table for fields to find min.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_apply_min_on">Field to sort for MIN.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to sort for MIN or not.</param>
	''' <returns>String</returns>
	Public Shared Function BuildMinString_GROUPED(t_ As String, field_to_group As String, field_to_apply_min_on As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT " & field_to_group & ", MIN(" & field_to_apply_min_on & ") "

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v & " GROUP BY " & field_to_group
	End Function

	''' <summary>
	''' Builds SQL MIN query string, grouped - useful for Pie chart.
	''' </summary>
	''' <param name="t_">Table for fields to find min.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_apply_min_on">Field to sort for MIN.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to sort for MIN or not, followed by the operator to apply on the key.</param>
	''' <returns>String</returns>
	Public Shared Function BuildMinString_GROUPED_CONDITIONAL(t_ As String, field_to_group As String, field_to_apply_min_on As String, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT " & field_to_group & ", MIN(" & field_to_apply_min_on & ") "

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v & " GROUP BY " & field_to_group
	End Function

	''' <summary>
	''' Builds SQL Select Top Query String.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
	''' <param name="OrderByField">Column to order by.</param>
	''' <param name="order_by">Ascending or Descending</param>
	''' <returns>String.</returns>
	Public Shared Function BuildTopString(t_ As String, Optional select_keys As Array = Nothing, Optional where_keys As Array = Nothing, Optional rows_to_select As Long = 1, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC) As String
		Dim v As String = "SELECT TOP " & Val(rows_to_select) & " * "
		If select_keys IsNot Nothing Then
			v = "SELECT TOP " & Val(rows_to_select) & " "
			With select_keys
				For i As Integer = 0 To .Length - 1
					v &= select_keys(i)
					If select_keys.Length > 1 And i <> select_keys.Length - 1 Then v &= ", "
				Next
			End With
			v &= " "
		End If
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If
		If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
		If OrderByField IsNot Nothing Then v &= " " & order_by.ToString

		Return v
	End Function

	Public Shared Function BuildTopString_GROUPED(t_ As String, field_to_group As String, Optional select_keys As Array = Nothing, Optional where_keys As Array = Nothing, Optional rows_to_select As Long = 10, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.DESC) As String
		Dim v As String = "SELECT TOP " & Val(rows_to_select) & " * "
		If select_keys IsNot Nothing Then
			v = "SELECT TOP " & Val(rows_to_select) & " "
			With select_keys
				For i As Integer = 0 To .Length - 1
					v &= select_keys(i)
					If select_keys.Length > 1 And i <> select_keys.Length - 1 Then v &= ", "
				Next
			End With
			v &= " "
		End If
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If
		If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
		If OrderByField IsNot Nothing Then v &= " " & order_by.ToString

		Return v & " GROUP BY " & field_to_group
	End Function

	Public Shared Function BuildTopString_CONDITIONAL(t_ As String, Optional select_keys As Array = Nothing, Optional where_key_operator As Array = Nothing, Optional rows_to_select As Long = 1, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC) As String
		Dim where_keys As Array = where_key_operator
		Dim select_params As Array = select_keys
		Dim v As String = "SELECT TOP " & Val(rows_to_select) & " * "
		If select_keys IsNot Nothing Then
			v = "SELECT TOP " & Val(rows_to_select) & " "
			With select_keys
				For i As Integer = 0 To .Length - 1
					v &= select_keys(i)
					If select_keys.Length > 1 And i <> select_keys.Length - 1 Then v &= ", "
				Next
			End With
			v &= " "
		End If
		v &= " FROM " & t_

		Dim l As New List(Of String)
		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					If l.Contains(where_keys(j)) Then
						v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j) & "_" & l.LastIndexOf(where_keys(j))
					Else
						v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					End If
					l.Add(where_keys(j))
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If
		If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
		If OrderByField IsNot Nothing Then v &= " " & order_by.ToString

		Return v
	End Function

	''' <summary>
	''' Builds SQL MAX query string, grouped - useful for Pie chart.
	''' </summary>
	''' <param name="t_">Table for fields to find MAX.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_apply_max_on">Field to sort for MAX.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to sort for MAX or not.</param>
	''' <returns>String</returns>
	Public Shared Function BuildMaxString_GROUPED(t_ As String, field_to_group As String, field_to_apply_max_on As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT " & field_to_group & ", MAX(" & field_to_apply_max_on & ") "
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v & " GROUP BY " & field_to_group
	End Function

	''' <summary>
	''' Builds SQL MAX query string, grouped - useful for Pie chart.
	''' </summary>
	''' <param name="t_">Table for fields to find MAX.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_apply_max_on">Field to sort for MAX.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to sort for MAX or not, followed by the operator to apply on the key.</param>
	''' <returns>String</returns>
	Public Shared Function BuildMaxString_GROUPED_CONDITIONAL(t_ As String, field_to_group As String, field_to_apply_max_on As String, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT " & field_to_group & ", MAX(" & field_to_apply_max_on & ") "
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v & " GROUP BY " & field_to_group
	End Function

	Public Shared Function BuildMaxString(t_ As String, Max_Field As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT MAX (" & Max_Field & ")"

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL AVG query string, ungrouped - useful for Line chart.
	''' </summary>
	''' <param name="t_">Table for fields to find AVG.</param>
	''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to sort for AVG or not.</param>
	''' <returns>String</returns>
	Public Shared Function BuildAVGString_UNGROUPED(t_ As String, field_to_apply_avg_on As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT AVG(" & field_to_apply_avg_on & ")"

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function


	''' <summary>
	''' Builds SQL AVG query string, ungrouped - useful for Line chart.
	''' </summary>
	''' <param name="t_">Table for fields to find AVG.</param>
	''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to sort for AVG or not, followed by the operator to apply on the key.</param>
	''' <returns>String</returns>
	Public Shared Function BuildAVGString_UNGROUPED_CONDITIONAL(t_ As String, field_to_apply_avg_on As String, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT AVG(" & field_to_apply_avg_on & ")"
		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function
	''' <summary>
	''' Builds SQL Sum query string, ungrouped - useful for Line chart.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="field_to_sum">Field to sum.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to sum or not.</param>
	''' <returns>String</returns>
	Public Shared Function BuildSumString_UNGROUPED(t_ As String, field_to_sum As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT SUM(" & field_to_sum & ") "

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL Sum query string, ungrouped - useful for Line chart.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="field_to_sum">Field to sum.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to sum or not, followed by the operator to apply on the key.</param>
	''' <returns>String</returns>
	Public Shared Function BuildSumString_UNGROUPED_CONDITIONAL(t_ As String, field_to_sum As String, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT SUM(" & field_to_sum & ") "

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function
	''' <summary>
	''' Builds SQL Count query string, ungrouped - useful for Line chart.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="where_keys">Fields to check equality condition on, to decide to count or not.</param>
	''' <param name="field_to_count">Field to count and group.</param>
	''' <returns>String</returns>
	Public Shared Function BuildCountString_UNGROUPED(t_ As String, field_to_count As String, Optional where_keys As Array = Nothing) As String
		Dim v As String = "SELECT COUNT(" & field_to_count & ")"

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1
					v &= where_keys(j) & "=@" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function


	''' <summary>
	''' Builds SQL Count query string, ungrouped - useful for Line chart.
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="where_key_operator">Fields to check condition on, to decide to count or not, followed by the operator to apply on the key.</param>
	''' <param name="field_to_count">Field to count and group.</param>
	''' <returns>String</returns>
	Public Shared Function BuildCountString_UNGROUPED_CONDITIONAL(t_ As String, field_to_count As String, Optional where_key_operator As Array = Nothing) As String
		Dim where_keys As Array = where_key_operator
		Dim v As String = "SELECT COUNT(" & field_to_count & ")"

		v &= " FROM " & t_

		If where_keys IsNot Nothing Then
			If where_keys.Length > 0 Then
				v &= " WHERE ("
				For j As Integer = 0 To where_keys.Length - 1 Step 2
					v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
					If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
						v &= " AND "
					End If
				Next
			End If
			v &= ")"
		End If

		Return v
	End Function

	''' <summary>
	''' Builds SQL Select Query String with DISTINCT. Useful for Line chart.
	''' </summary>
	''' <param name="t_">Table to select from.</param>
	''' <param name="select_params">Fields to select.</param>
	''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO">Columns to consider with condition, appended with _FROM or _TO.</param>
	''' <param name="OrderByField">Column to order by.</param>
	''' <param name="order_by">Ascending or Descending</param>
	''' <returns>String.</returns>
	Public Shared Function BuildSelectString_DISTINCT_BETWEEN_CONDITIONAL(t_ As String, select_params As Array, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC) As String

		Dim v As String = "SELECT DISTINCT "

		If select_params IsNot Nothing Then
			For i As Integer = 0 To select_params.Length - 1
				v &= select_params(i)
				If select_params.Length > 1 And i <> select_params.Length - 1 Then
					v &= ", "
				End If
			Next
		Else
			v &= " *"
		End If
		v &= " FROM " & t_

		If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
			If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
				v &= " WHERE "
				For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
					v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO)"
					If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If

		End If
		If OrderByField IsNot Nothing Then
			v &= " ORDER BY " & OrderByField
		End If
		If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
		Return v
	End Function
	''' <summary>
	''' Builds SQL AVG query string, grouped - useful for Line chart.
	''' </summary>
	''' <param name="t_">Table for fields to find AVG.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
	''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO"></param>
	''' <returns>String</returns>
	Public Shared Function BuildAVGString_GROUPED_BETWEEN(t_ As String, field_to_group As String, field_to_apply_avg_on As String, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing) As String

		Dim v As String = "SELECT " & field_to_group & ", AVG(" & field_to_apply_avg_on & ") "
		v &= " FROM " & t_

		If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
			If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
				v &= " WHERE "
				For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
					v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO)"
					If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If
		End If
		Return v & " GROUP BY " & field_to_group
	End Function
	''' <summary>
	''' Builds SQL SUM query string, grouped - useful for Line chart.
	''' </summary>
	''' <param name="t_">Table for fields to find AVG.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_sum">Field to sum.</param>
	''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO"></param>
	''' <returns>String</returns>
	Public Shared Function BuildSumString_GROUPED_BETWEEN(t_ As String, field_to_group As String, field_to_sum As String, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing) As String

		Dim v As String = "SELECT " & field_to_group & ", SUM(" & field_to_sum & ") "
		v &= " FROM " & t_

		If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
			If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
				v &= " WHERE "
				For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
					v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO)"
					If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If

		End If
		Return v & " GROUP BY " & field_to_group
	End Function
	''' <summary>
	''' Builds SQL Count query string, grouped - useful for Line chart. When declaring keys, field must not contain _From or _To. When declaring keys_values, field must have _From and _To, e.g. RecordDate_From, RecordDate_To
	''' </summary>
	''' <param name="t_">Table for fields to count.</param>
	''' <param name="field_to_group">Field to group.</param>
	''' <param name="field_to_count">Field to count.</param>
	''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO">single field name, e.g. RecordDate</param>
	''' <returns>String</returns>
	Public Shared Function BuildCountString_GROUPED_BETWEEN(t_ As String, field_to_group As String, field_to_count As String, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing) As String

		Dim v As String = "SELECT " & field_to_group & ", COUNT(" & field_to_count & ") "
		v &= " FROM " & t_

		If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
			If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
				v &= " WHERE "
				For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
					v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO)"
					If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
						v &= " AND "
					End If
				Next
			End If

		End If
		Return v & " GROUP BY " & field_to_group
	End Function


#End Region

#Region "CSV"
	Public Shared Sub CommitData(where_to_place_data As CommitDataTargetInfo, query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing)
		If Mid(query, 1, Len("select")).ToLower <> "select" And Mid(query, 1, Len("update")).ToLower <> "update" Then
			Return
		End If

		Dim data As DataTable = GetDataTable(query, connection_string, parameters_keys_values_)
		Dim header As String = ""
		Dim content As String = ""

		With data
			'header
			For h = 0 To .Columns.Count - 1
				header &= .Columns(h).ColumnName
				If h <> .Columns.Count Then header &= ","
			Next

			For i = 0 To .Rows.Count - 1
				For j = 0 To .Columns.Count - 1
					'content
					content &= .Rows(i).Item(j)
					If j <> .Columns.Count Then content &= ","
				Next
				content &= vbCrLf
			Next
		End With

		WriteText(where_to_place_data.filename, header.Trim & vbCrLf & content.Trim)
	End Sub

#End Region

#Region "Retrieval"
	Public Shared Function GetDataTable(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As DataTable

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

			'da = New SqlDataAdapter(Command)
			'dt = New DataTable
			da.Fill(dt)
			Return dt
		Catch
			Throw New Exception("Table is empty")
		End Try

	End Function


	Public Enum Optimization
		FaultTolerant = 0
		AsIs = 1
	End Enum
	''' <summary>
	''' Executes SQL statement for a single value.
	''' </summary>
	''' <param name="query_">SQL Query. You can use BuildSelectString, BuildUpdateString, BuildInsertString, BuildCountString, BuildTopString instead.</param>
	''' <param name="connection_string">Connection String.</param>
	''' <param name="parameters_keys_values_">Parameters.</param>
	''' <param name="return_type_is">Return string by default.</param>
	''' <returns></returns>
	Public Shared Function QData(query_ As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing, Optional return_type_is As Optimization = Optimization.FaultTolerant)
		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = parameters_keys_values_

		Dim con As SqlConnection = New SqlConnection(connection_string)
		Dim cmd As SqlCommand = New SqlCommand(query_, con)

		If select_parameter_keys_values IsNot Nothing Then
			For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
				cmd.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
			Next
		End If

		Try
			Using con
				con.Open()
				If return_type_is = Optimization.AsIs Then
					'					Return CType(cmd.ExecuteScalar(), String).ToString
					Return cmd.ExecuteScalar()
				ElseIf return_type_is = Optimization.FaultTolerant Then
					Return CType(cmd.ExecuteScalar(), String).ToString
				End If
			End Using
		Catch
		End Try

	End Function

	Public Shared Function QList(query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As List(Of String)
		Dim l As New List(Of String)

		Dim select_parameter_keys_values() = {}
		select_parameter_keys_values = select_parameter_keys_values_

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

			With dt
				For i = 0 To .Rows.Count - 1
					l.Add(.Rows(i).Item(0).ToString)
				Next
			End With
		Catch
		End Try

		Return l
	End Function

	'Public Shared Function QDataWORKING(query_ As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing, Optional return_type_is As Want = Want.Value)
	'	Dim select_parameter_keys_values() = {}
	'	select_parameter_keys_values = parameters_keys_values_

	'	Dim con As SqlConnection = New SqlConnection(connection_string)
	'	Dim cmd As SqlCommand = New SqlCommand(query_, con)

	'	If select_parameter_keys_values IsNot Nothing Then
	'		For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
	'			cmd.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
	'		Next
	'	End If

	'	Try
	'		Using con
	'			con.Open()
	'			If return_type_is = Want.Value Then
	'				Return CType(cmd.ExecuteScalar(), String).ToString
	'			ElseIf return_type_is = Want.Display Then
	'			End If
	'		End Using
	'	Catch
	'	End Try

	'End Function

	''' <summary>
	''' Checks if a field exists w/ or w/o specified condition.
	''' </summary>
	''' <param name="t_">Table to perform operation on.</param>
	''' <param name="connection_string">Connection String.</param>
	''' <param name="where_keys">List of parameters.</param>
	''' <param name="where_keys_values">List of parameters and their values.</param>
	''' <returns>True if it exists, False otherwise.</returns>

	Public Shared Function QExists(t_ As String, connection_string As String, Optional where_keys As Array = Nothing, Optional where_keys_values As Array = Nothing) As Boolean
		Return QCount(t_, connection_string, where_keys, where_keys_values) > 0
	End Function

	Public Shared Function QExists_CONDITIONAL(t_ As String, connection_string As String, Optional where_key_operator As Array = Nothing, Optional where_keys_values As Array = Nothing) As Boolean
		Return QData(BuildCountString_CONDITIONAL(t_, where_key_operator), connection_string, where_keys_values) > 0
	End Function

	''' <summary>
	''' Highest value of Max_Field for given choices w/ or w/o specified condition.
	''' </summary>
	''' <param name="t_">Table to perform operation on.</param>
	''' <param name="connection_string">Connection String.</param>
	''' <param name="where_keys">List of parameters.</param>
	''' <param name="where_keys_values">List of parameters and their values.</param>
	''' <param name="Max_Field">Field to use as maximum.</param>
	''' <returns>Value of Max_Field.</returns>
	Public Shared Function QMax(t_ As String, connection_string As String, Optional where_keys As Array = Nothing, Optional where_keys_values As Array = Nothing, Optional Max_Field As String = Nothing)
		Return QData(BuildMaxString(t_, Max_Field, where_keys), connection_string, where_keys_values)
	End Function

	Public Shared Function QCount(t_ As String, connection_string As String, Optional where_keys As Array = Nothing, Optional where_keys_values As Array = Nothing)
		Return QData(BuildCountString(t_, where_keys), connection_string, where_keys_values)
	End Function
	Public Shared Function QCount_CONDITIONAL(t_ As String, connection_string As String, Optional where_key_operator As Array = Nothing, Optional where_keys_values As Array = Nothing)
		Return QData(BuildCountString_CONDITIONAL(t_, where_key_operator), connection_string, where_keys_values_CONDITIONAL(where_keys_values))

	End Function
#End Region

#Region "Retrieval - Raw"
	Public Shared Function QTable(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As DataTable

		Try
			Using con As New SqlConnection(connection_string)
				Using cmd As New SqlCommand(query)
					Using sda As New SqlDataAdapter()
						cmd.Connection = con
						If select_parameter_keys_values IsNot Nothing Then
							For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
								cmd.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
							Next
						End If
						sda.SelectCommand = cmd
						Using dt As New DataTable
							sda.Fill(dt)
							Return dt
						End Using
					End Using
				End Using
			End Using
		Catch ex As Exception
		End Try

	End Function

#End Region

#Region "Code"
	Public Structure QParameters
		Public operation As Queries
		Public table As String
		Public SelectColumns As Array
		Public WhereKeys As Array
		Public InsertKeys As Array
		Public WhereOperators As Array
		Public MaxField As String
		Public OrderByField As String
		Public OrderRecordsBy As OrderBy
		Public LikeSelect As LIKE_SELECT
		Public TopRowsToSelect As Long
		Public UpdateKeys As Array
	End Structure
	Public Enum QOutput
		QData
		Display
		QString
		Commit
		QExists
		QCount
		QCount_Conditional
		BindProperty
	End Enum
	Public Enum Output
		Web
		Desktop
	End Enum

	Public Shared Function QString(q_parameters As QParameters, Optional q_output As QOutput = QOutput.QData, Optional connection_string_code_variable As String = Nothing, Optional output_ As Output = Output.Web) As String
		Dim operation As Queries = q_parameters.operation, table As String = q_parameters.table, SelectColumns As Array = q_parameters.SelectColumns, WhereKeys As Array = q_parameters.WhereKeys, InsertKeys As Array = q_parameters.InsertKeys, WhereOperators As Array = q_parameters.WhereOperators, MaxField As String = q_parameters.MaxField, OrderByField As String = q_parameters.OrderByField, OrderRecordsBy As OrderBy = q_parameters.OrderRecordsBy, LikeSelect As LIKE_SELECT = q_parameters.LikeSelect, TopRowsToSelect As Long = q_parameters.TopRowsToSelect, UpdateKeys As Array = q_parameters.UpdateKeys
		Dim r As String

		If operation = Queries.BuildUpdateString Then
			'BuildUpdateString(table__, {uk}, {wk})
			r = "BuildUpdateString(""" & table & """" 'WhereKeys)"
			If UpdateKeys IsNot Nothing Then
				With UpdateKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & UpdateKeys(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If WhereKeys IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			r &= ")"
		End If

		If operation = Queries.BuildTopString Then
			'BuildTopString(table, {sk}, {wk}, rows_to_select:=1, "order_by_field", OrderBy.ASC)
			r = "BuildTopString(""" & table & """" 'WhereKeys)"
			If SelectColumns IsNot Nothing Then
				With SelectColumns
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & SelectColumns(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If WhereKeys IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			r &= ", " & TopRowsToSelect
			If OrderByField IsNot Nothing Then
				r &= ", """ & OrderByField & """"
				r &= ", OrderBy." & OrderRecordsBy.ToString
			End If
			r &= ")"
		End If

		If operation = Queries.BuildTopString_CONDITIONAL Then
			'BuildTopString_CONDITIONAL(table, {sk}, {wko}, rows_to_select:=1, "order_by_field", OrderBy.ASC)
			r = "BuildTopString_CONDITIONAL(""" & table & """" 'WhereKeys)"
			If SelectColumns IsNot Nothing Then
				With SelectColumns
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & SelectColumns(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If WhereKeys IsNot Nothing And WhereOperators IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """, """ & WhereOperators(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If TopRowsToSelect > 0 Then
				r &= ", " & TopRowsToSelect
			Else
				r &= ", 1"
			End If
			If OrderByField IsNot Nothing Then
				r &= ", """ & OrderByField & """"
				r &= ", OrderBy." & OrderRecordsBy.ToString
			End If
			r &= ")"
		End If

		If operation = Queries.BuildSelectString_LIKE Then
			'BuildSelectString_LIKE(table__, {sk}, {wk}, "order_by_field", OrderBy.ASC, LIKE_SELECT.AND_)
			r = "BuildSelectString_LIKE(""" & table & """" 'WhereKeys)"
			If SelectColumns IsNot Nothing Then
				With SelectColumns
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & SelectColumns(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If WhereKeys IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If OrderByField IsNot Nothing Then
				r &= ", """ & OrderByField & """"
				r &= ", OrderBy." & OrderRecordsBy.ToString
			End If
			If LikeSelect <> Nothing Then
				r &= ", LIKE_SELECT." & LikeSelect.ToString
			End If
			r &= ")"
		End If

		If operation = Queries.BuildSelectString_DISTINCT Then
			'BuildSelectString_DISTINCT(table__, {SelectColumns}, {wk})
			r = "BuildSelectString_DISTINCT(""" & table & """" 'WhereKeys)"
			If SelectColumns IsNot Nothing Then
				With SelectColumns
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & SelectColumns(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If WhereKeys IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			r &= ")"
		End If

		If operation = Queries.BuildSelectString_CONDITIONAL Then
			'BuildSelectString_CONDITIONAL(table__, {SelectColumns}, {wko}, "order_by_field", OrderBy.ASC)
			r = "BuildSelectString_CONDITIONAL(""" & table & """" 'WhereKeys)"
			If SelectColumns IsNot Nothing Then
				With SelectColumns
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & SelectColumns(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If WhereKeys IsNot Nothing And WhereOperators IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """, """ & WhereOperators(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If OrderByField IsNot Nothing Then
				r &= ", """ & OrderByField & """"
				r &= ", OrderBy." & OrderRecordsBy.ToString
			End If
			r &= ")"
		End If

		If operation = Queries.BuildSelectString_BETWEEN Then
			'BuildSelectString_BETWEEN(table__, {SelectColumns}, {WhereKeys}, "order_by_field", OrderBy.ASC)
			'SELECT MyAdminMedia FROM MyAdminMedia WHERE (RecordSerial BETWEEN @RecordSerial_FROM AND @RecordSerial_TO)
			r = "BuildSelectString_BETWEEN(""" & table & """" 'WhereKeys)"
			If SelectColumns IsNot Nothing Then
				With SelectColumns
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & SelectColumns(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If WhereKeys IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If OrderByField IsNot Nothing Then
				r &= ", """ & OrderByField & """"
				r &= ", OrderBy." & OrderRecordsBy.ToString
			End If
			r &= ")"
		End If

		If operation = Queries.BuildSelectString Then
			'BuildSelectString(table__, {SelectColumns}, {WhereKeys}, "order_by_field", OrderBy.ASC)
			r = "BuildSelectString(""" & table & """" 'WhereKeys)"
			If SelectColumns IsNot Nothing Then
				With SelectColumns
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & SelectColumns(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If WhereKeys IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			Else
				r &= ", Nothing"
			End If
			If OrderByField IsNot Nothing Then
				r &= ", """ & OrderByField & """"
				r &= ", OrderBy." & OrderRecordsBy.ToString
			End If
			r &= ")"
		End If

		If operation = Queries.BuildMaxString Then
			'BuildMaxString(table, where_keys, Max_Field)
			r = "BuildMaxString(""" & table & """" 'WhereKeys)"
			If WhereKeys IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}, "
				End With
			Else
				r &= ", Nothing, "
			End If
			r &= """" & MaxField & """)"
		End If

		If operation = Queries.BuildInsertString Then
			'BuildInsertString(table, InsertKeys)
			r = "BuildInsertString(""" & table & """" 'WhereKeys)"
			If InsertKeys IsNot Nothing Then
				With InsertKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & InsertKeys(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			End If
			r &= ")"
		End If

		If operation = Queries.BuildCountString_CONDITIONAL Then
			'BuildCountString_CONDITIONAL(table, WhereKeyOperator)
			r = "BuildCountString_CONDITIONAL(""" & table & """" 'WhereKeys)"
			If WhereKeys IsNot Nothing And WhereOperators IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """, """ & WhereOperators(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			End If
			r &= ")"
		End If

		If operation = Queries.BuildCountString Then
			'BuildCountString(table, WhereKeys)
			r = "BuildCountString(""" & table & """" 'WhereKeys)"
			If WhereKeys IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			End If
			r &= ")"
		End If

		If operation = Queries.DeleteString_CONDITIONAL Then
			'DeleteString_CONDITIONAL(table, {wko})
			r = "DeleteString_CONDITIONAL(""" & table & """" 'WhereKeys)"

			If WhereKeys IsNot Nothing And WhereOperators IsNot Nothing Then
				With WhereKeys
					r &= ", {"
					For i As Integer = 0 To .Length - 1
						r &= """" & WhereKeys(i) & """, """ & WhereOperators(i) & """"
						If i < .Length - 1 Then r &= ", "
					Next
					r &= "}"
				End With
			End If
			r &= ")"
		End If

		If q_output = QOutput.QString Then Return r

		Dim wko As String = ""
		If WhereKeys IsNot Nothing And WhereOperators IsNot Nothing Then
			With WhereKeys
				wko = "{"
				For i As Integer = 0 To .Length - 1
					wko &= """" & WhereKeys(i) & """, """ & WhereOperators(i) & """"
					If i < .Length - 1 Then
						wko &= ", "
					End If
				Next
				wko &= "}"
			End With
		End If

		Dim wk As String = ""
		If WhereKeys IsNot Nothing Then
			With WhereKeys
				wk = "{"
				For i As Integer = 0 To .Length - 1
					wk &= """" & WhereKeys(i) & """"
					If i < .Length - 1 Then
						wk &= ", "
					End If
				Next
				wk &= "}"
			End With
		End If

		Dim wkv_b As String = ""
		If WhereKeys IsNot Nothing Then
			With WhereKeys
				wkv_b = "{"
				For i As Integer = 0 To .Length - 1
					wkv_b &= """" & WhereKeys(i) & "_FROM"", """", """ & WhereKeys(i) & "_TO"", """""
					If i < .Length - 1 Then
						wkv_b &= ", "
					End If
				Next
				wkv_b &= "}"
			End With
		End If

		Dim wkv As String = ""
		Dim l_wkv_conditional As New List(Of String)
		If WhereKeys IsNot Nothing Then
			If operation = Queries.BuildSelectString_CONDITIONAL Or operation = Queries.BuildCountString_CONDITIONAL Or operation = Queries.BuildTopString_CONDITIONAL Then
				With WhereKeys
					wkv = "{"
					For i As Integer = 0 To .Length - 1
						If l_wkv_conditional.Contains(WhereKeys(i)) Then
							wkv &= """" & WhereKeys(i) & "_" & l_wkv_conditional.LastIndexOf(WhereKeys(i)) & """, """" "
						Else
							wkv &= """" & WhereKeys(i) & """, """" "
						End If
						l_wkv_conditional.Add(WhereKeys(i))
						If i < .Length - 1 Then
							wkv &= ", "
						End If
					Next
					wkv &= "}"
				End With
			Else
				With WhereKeys
					wkv = "{"
					For i As Integer = 0 To .Length - 1
						wkv &= """" & WhereKeys(i) & """, """" "
						If i < .Length - 1 Then
							wkv &= ", "
						End If
					Next
					wkv &= "}"
				End With
			End If
		End If

		Dim ikv As String = ""
		If InsertKeys IsNot Nothing Then
			With InsertKeys
				ikv = "{"
				For i As Integer = 0 To .Length - 1
					ikv &= """" & InsertKeys(i) & """, """" "
					If i < .Length - 1 Then
						ikv &= ", "
					End If
				Next
				ikv &= "}"
			End With
		End If

		Dim ukv As String = ""
		If UpdateKeys IsNot Nothing Then
			With UpdateKeys
				ukv = "{"
				For i As Integer = 0 To .Length - 1
					ukv &= """" & UpdateKeys(i) & """, """" "
					If i < .Length - 1 Then
						ukv &= ", "
					End If
				Next
				If WhereKeys IsNot Nothing Then
					With WhereKeys
						ukv &= ", "
						For i As Integer = 0 To .Length - 1
							ukv &= """" & WhereKeys(i) & """, """" "
							If i < .Length - 1 Then
								ukv &= ", "
							End If
						Next
					End With
				End If
				ukv &= "}"
			End With
		End If

		Dim q As String = ""

		If q_output = QOutput.BindProperty Then
			If output_ = Output.Desktop Then
				'BindProperty(New Object, PropertyToBind.Items, r, "con", {kv}, "dataTextField", "dataValueField")
				q = "BindProperty(control_here, PropertyToBind.Items, " & r & ", connection_string_here"
				If connection_string_code_variable IsNot Nothing Then q = "BindProperty(control_here, PropertyToBind.Items, " & r & ", " & connection_string_code_variable
				If wkv.Length > 0 Then
					q &= ", " & wkv
				Else
					q &= ", Nothing"
				End If
				q &= ", """ & SelectColumns(0) & """"
				q &= ")"
			Else
				'BindProperty(control_here, r, "con", {kv}, "dtf")
				q = "BindProperty(control_here, " & r & ", connection_string_here"
				If connection_string_code_variable IsNot Nothing Then q = "BindProperty(control_here, " & r & ", " & connection_string_code_variable
				If wkv.Length > 0 Then
					q &= ", " & wkv
				Else
					q &= ", Nothing"
				End If
				q &= ", """ & SelectColumns(0) & """"
				q &= ")"
			End If
			Return q
		End If

		If q_output = QOutput.QCount Then
			'QCount(table, "conn", {where_keys()}, {wherekv})
			q = "QCount(""" & table & """, connection_string_here"
			If connection_string_code_variable IsNot Nothing Then q = "QCount(""" & table & """, " & connection_string_code_variable

			If wk.Length > 0 Then
				q &= ", " & wk
			End If
			If wkv.Length > 0 Then
				q &= ", " & wkv
			End If
			q &= ")"
			Return q
		End If

		If q_output = QOutput.QCount_Conditional Then
			'QCount_CONDITIONAL(table, "con", {wko}, {wkv})
			q = "QCount_CONDITIONAL(""" & table & """, connection_string_here"
			If connection_string_code_variable IsNot Nothing Then q = "QCount_CONDITIONAL(""" & table & """, " & connection_string_code_variable

			If wko.Length > 0 Then
				q &= ", " & wko
			End If
			If wkv.Length > 0 Then
				q &= ", " & wkv
			End If
			q &= ")"
			Return q
		End If

		If q_output = QOutput.QExists Then
			'QExists(table, server_con, {wk}, {wkv})
			q = "QExists(""" & table & """, connection_string_here"
			If connection_string_code_variable IsNot Nothing Then q = "QExists(""" & table & """, " & connection_string_code_variable

			If wk.Length > 0 Then
				q &= ", " & wk
			End If
			If wkv.Length > 0 Then
				q &= ", " & wkv
			End If
			q &= ")"
			Return q
		End If

		If q_output = QOutput.QData Then
			'QData(r, server_con, {KV})
			q = "QData(" & r & ", connection_string_here"
			If connection_string_code_variable IsNot Nothing Then q = "QData(" & r & ", " & connection_string_code_variable
			If WhereKeys IsNot Nothing Then
				If operation <> Queries.BuildSelectString_BETWEEN Then
					q &= ", " & wkv
				Else
					q &= ", " & wkv_b
				End If
			End If
			q &= ")"
			Return q
		End If

		If q_output = QOutput.Display Then
			'Display(grid, q)
			q = "Display(grid_here, " & r & ", connection_string_here"
			If connection_string_code_variable IsNot Nothing Then q = "Display(grid_here, " & r & ", " & connection_string_code_variable
			If WhereKeys IsNot Nothing Then
				If operation <> Queries.BuildSelectString_BETWEEN Then
					q &= ", " & wkv
				Else
					q &= ", " & wkv_b
				End If
			End If
			q &= ")"
			Return q
		End If

		If q_output = QOutput.Commit Then
			q = "CommitSequel(" & r & ", connection_string_here"
			If connection_string_code_variable IsNot Nothing Then q = "CommitSequel(" & r & ", " & connection_string_code_variable
			'If WhereKeys IsNot Nothing Then
			If operation = Queries.BuildInsertString Then
				q &= ", " & ikv
			ElseIf operation = Queries.BuildUpdateString Then
				q &= ", " & ukv
			ElseIf operation = Queries.BuildUpdateString_CONDITIONAL Then
				q &= ", " & ukv
			End If
			q &= ")"
			Return q
		End If


	End Function

#End Region

#Region "Support Functions"
	Private Shared Function where_keys_values_CONDITIONAL(where_keys_values As Array) As Array
		Dim k As Array = where_keys(where_keys_values)
		Dim v As Array = where_values(where_keys_values)
		Dim l As New List(Of String)
		Dim lkv As New List(Of String)
		With k
			For i As Integer = 0 To .Length - 1
				If l.Contains(k(i)) Then
					lkv.Add(k(i) & "_" & l.LastIndexOf(k(i)))
					lkv.Add(v(i))
				Else
					lkv.Add(k(i))
					lkv.Add(v(i))
				End If
				l.Add(k(i))
			Next
		End With

		Return lkv.ToArray
	End Function
	Private Shared Function where_keys_emptyValues_CONDITIONAL(where_keys As Array) As Array
		Dim k As Array = where_keys(where_keys)
		Dim l As New List(Of String)
		Dim lkv As New List(Of String)
		With k
			For i As Integer = 0 To .Length - 1
				If l.Contains(k(i)) Then
					lkv.Add(k(i) & "_" & l.LastIndexOf(k(i)))
					lkv.Add("""""")
				Else
					lkv.Add(k(i))
					lkv.Add("""""")
				End If
				l.Add(k(i))
			Next
		End With

		Return lkv.ToArray
	End Function

	Public Shared Function where_keys(where_keys_values As Array) As Array
		If where_keys_values Is Nothing Then Return Nothing
		If where_keys_values.Length < 1 Then Return Nothing
		Dim l As New List(Of String)
		With where_keys_values
			For i As Integer = 0 To .Length - 2 Step 2
				l.Add(where_keys_values(i))
			Next
		End With
		Return l.ToArray
	End Function

	Public Shared Function where_values(where_keys_values As Array) As Array
		If where_keys_values Is Nothing Then Return Nothing
		If where_keys_values.Length < 1 Then Return Nothing
		Dim l As New List(Of String)
		With where_keys_values
			For i As Integer = 0 To .Length - 2 Step 2
				l.Add(where_keys_values(i + 1))
			Next
		End With
		Return l.ToArray
	End Function
#End Region

#Region "Checks"
	Public Shared Function TableHasData(table_ As String, con_string As String) As Boolean
		Return QExists(table_, con_string)
	End Function
	Public Shared Function TableExists(table_ As String, con_string As String) As Boolean

		Return QList(BuildSelectString("sys.Tables", {"name"}, Nothing, "name"), con_string).Contains(table_)

	End Function
#End Region

End Class
