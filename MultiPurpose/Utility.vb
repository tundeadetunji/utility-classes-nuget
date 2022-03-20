Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.AccessControl
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Public Class Utility


#Region ""
    Enum FormatFor
        Web
        JavaScript
        Custom
    End Enum
    Public Enum AppWindowState
        Hide = 0
        ShowNormal = 1
        ShowMinimized = 2
        ShowMaximized = 3
        ShowNormalNoActivate = 4
        Show = 5
        Minimize = 6
        ShowMinNoActivate = 7
        ShowNoActivate = 8
        Restore = 9
        ShowDefault = 10
        ForceMinimized = 11
    End Enum

    Public Enum DropTextPattern As Byte
        AlwaysNever
        YesNo
        OnOff
        OneZero
        TrueFalse
        IfNot
        SucceededFailed
        Cart
    End Enum
    Public Enum ListIsOf
        String_
        Object_
        Integer_
        Double_
        Short_
        Byte_

    End Enum

    Public Enum TextCase
        Capitalize = 0
        UpperCase = 1
        LowerCase = 2
        None = 3
    End Enum

    Public Enum SideToReturn
        Left
        Right
        AsArray
        AsListOfString
        AsListToString
    End Enum
    Public Enum IDPattern
        Short_
        Short_DateOnly
        Short_DateTime
        Long_
        Long_DateTime

    End Enum
    Public Enum SingularWord
        Message
        Mark
        Minute
        Day
        Student
        Month
        Year
        Hour
        Second
        Account
        Male
        Female
        User
        File
        Record
        Question
        Task
        Item
        Unit
        Room
        Client
        Match
        Invoice
        Sale
        Receipt
        Product
    End Enum

#End Region
    Public Shared Function GetFiles(directory_ As String, Optional ext_ As String = "*.txt", Optional search_depth As FileIO.SearchOption = FileIO.SearchOption.SearchTopLevelOnly) As ReadOnlyCollection(Of String)
        Dim Files__ As ReadOnlyCollection(Of String)
        Try
            Files__ = My.Computer.FileSystem.GetFiles(directory_, search_depth, ext_)
            If Files__.Count > 0 Then
                Return Files__
            Else
                Return Nothing
            End If
        Catch
        End Try
    End Function

    ''' <summary>
    ''' Formats to 2 decimal places. Same as RoundNumber(val)
    ''' </summary>
    ''' <param name="val_"></param>
    ''' <returns></returns>
    Public Shared Function ToCurrency(val_)
        Try
            Return RoundNumber(Val(val_))
        Catch
        End Try
    End Function
    ''' <summary>
    ''' Formats to 2 decimal places.
    ''' </summary>
    ''' <param name="val_"></param>
    ''' <returns></returns>
    Public Shared Function RoundNumber(val_)
        Dim val__
        If Val(val_) = 0 Then
            val__ = 0
        Else
            val__ = val_
        End If
        Dim return_ = FormatNumber(val__, 2, TriState.False, TriState.False, TriState.False)
        If return_.ToString = ".00" Then return_ = "0.00"
        Return return_
    End Function


    ''' <summary>
    ''' Strips HTML document of all HTML tags. Same as HTMLToText().
    ''' </summary>
    ''' <param name="html_markup"></param>
    ''' <returns></returns>
    Public Shared Function RemoveHTMLFromText(ByVal html_markup As String) As String
        Dim rx As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("<[^>]*>")
        Return rx.Replace(html_markup, "")
    End Function

    ''' <summary>
    ''' Strips HTML document of all HTML tags.
    ''' </summary>
    ''' <param name="html_markup"></param>
    ''' <returns></returns>
    Public Shared Function HTMLToText(ByVal html_markup As String) As String
        Return RemoveHTMLFromText(html_markup)
    End Function

    ''' <summary>
    ''' Adds a random number to the list, keeping all the list's items unique.
    ''' </summary>
    ''' <param name="random_inclusive_min">Where to start (used by random)</param>
    ''' <param name="random_exclusive_max">Where to stop (used by random)</param>
    ''' <param name="already_">The list that to add to</param>
    ''' <returns>The list with the new number</returns>

    Public Shared Function RandomList(random_inclusive_min As Integer, random_exclusive_max As Integer, already_ As List(Of Integer), Optional recycle_ As Boolean = True) As Integer
        Dim r_val
        If already_.Count >= random_exclusive_max Then
            If recycle_ = True Then
                already_.Clear()
            Else
                Return already_(already_.Count - 1)
            End If
        End If
2:
        Try
            r_val = Random_(random_inclusive_min, random_exclusive_max)
            If already_.Count > 0 And already_.Contains(r_val) Then
                GoTo 2
            Else
                already_.Add(r_val)
                Return r_val
            End If
        Catch
        End Try
    End Function

    ''' <summary>
    ''' Creates a shuffled (integer) list.
    ''' </summary>
    ''' <param name="items_from"></param>
    ''' <param name="items_to"></param>
    ''' <param name="items_"></param>
    ''' <returns></returns>
    Public Function CreateRandomList(items_from As Integer, items_to As Integer, items_ As List(Of Integer)) As List(Of Integer)
        For i As Integer = items_from To items_to
            RandomList(items_from, items_to + 1, items_)
        Next
        Return items_


        '		Dim temp_list As New List(Of Integer)
        '		g.CreateRandomList(1, 7, temp_list)

        '		'Then use the list, for example:
        '		If temp_list.Count < 1 Then Exit Sub
        '		With temp_list
        '			For j As Integer = 0 To .Count - 1
        '				u.Text &= .Item(j).ToString & vbCrLf
        '			Next
        '		End With

    End Function

    Public Shared Function Random_(inclusive_min As Integer, exclusive_max As Integer) As Integer
        Dim generator As New Random
        Dim randomValue As Integer
        randomValue = generator.Next(inclusive_min, exclusive_max)
        Return randomValue
    End Function


    ''' <summary>
    ''' Returns the opposite of the (boolean) value.
    ''' </summary>
    ''' <param name="boolean_val">True or False</param>
    ''' <returns>The opposite of boolean_val</returns>
    Public Shared Function BooleanOpposite(boolean_val As Boolean) As Boolean
        Select Case boolean_val
            Case True
                Return False
            Case False
                Return True
        End Select
    End Function

    ''' <summary>
    ''' Splits text to left and right sides of separator_ and returns either left or right side depending on side_to_return.
    ''' </summary>
    ''' <param name="string_to_split"></param>
    ''' <param name="separator">delimeter</param>
    ''' <param name="side_to_return">Left or Right</param>
    ''' <returns>String</returns>
    Public Shared Function SplitTextInTwo(string_to_split As String, separator As String, Optional side_to_return As SideToReturn = SideToReturn.Right)
        If string_to_split.Length < 1 Or separator.Length < 1 Then Return ""

        If side_to_return = SideToReturn.Left Then
            Return string_to_split.Split(separator, 2, StringSplitOptions.None)(0)
        Else
            Try
                Return string_to_split.Split(separator, 2, StringSplitOptions.None)(1)
            Catch ex As Exception
                Return ""
            End Try
        End If

    End Function

    ''' <summary>
    ''' Splits string into multiple parts.
    ''' </summary>
    ''' <param name="string_to_split"></param>
    ''' <param name="separator"></param>
    ''' <param name="side_to_return"></param>
    ''' <returns></returns>
    Public Shared Function SplitTextInSplits(string_to_split As String, separator As String, Optional side_to_return As SideToReturn = SideToReturn.AsArray)
        If string_to_split.Length < 1 Then Return Nothing

        Dim str As String
        Dim strArr() As String
        str = string_to_split
        strArr = str.Split(separator)

        Dim l As New List(Of String)

        Select Case side_to_return
            Case SideToReturn.AsArray
                Return strArr
            Case SideToReturn.AsListOfString
                With strArr
                    For i = 0 To .Length - 1
                        l.Add(strArr(i))
                    Next
                End With
                Return l
            Case SideToReturn.AsListToString
                Return ListToString(strArr)
        End Select
    End Function
    Private Shared Function LeadingZero(number As String) As String
        If number.Length < 2 Then Return "0" & number Else : Return number
    End Function

    ''' <summary>
    ''' Returns a unique ID, depending on the pattern.
    ''' </summary>
    ''' <param name="case_acct_or_date_time"></param>
    ''' <param name="replace_guid"></param>
    ''' <returns></returns>
    Public Shared Function NewID(Optional case_acct_or_date_time As String = "date_time", Optional replace_guid As String = "") As String
        Dim raw_id As String = System.Guid.NewGuid().ToString, counter As Integer

        If replace_guid.Trim.Length > 0 Then
            raw_id = replace_guid.Trim
            GoTo 2
        End If

        For i% = 1 To raw_id.Length
            If Mid(raw_id, i, 1) = "-" Then
                counter += 1
                If counter = 2 Then
                    raw_id = Mid(raw_id, 1, i - 1)
                ElseIf counter = 0 Then
                    Exit For
                End If
            End If
        Next

2:
        Select Case case_acct_or_date_time.ToLower
            Case "acct"
                raw_id = raw_id & "-" & My.Computer.Clock.LocalTime.Date.ToShortDateString
            Case "date_time"
                raw_id = raw_id & "-" & Now.Year.ToString & "." & Now.Month.ToString & "." & Now.Day.ToString & "-" & Now.Hour.ToString & "." & LeadingZero(Now.Minute.ToString)
            Case ""
        End Select
        Return raw_id
    End Function
    ''' <summary>
    ''' Returns a unique ID, depending on the pattern.
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function NewID(IDPattern_ As IDPattern, Optional prefx As String = "") As String
        Select Case IDPattern_
            Case IDPattern.Long_
                Return NewGUID(prefx)
            Case IDPattern.Long_DateTime
                Return NewGUID(prefx, True)
            Case IDPattern.Short_
                Return NewID("", prefx)
            Case IDPattern.Short_DateOnly
                Return NewID("acct")
            Case IDPattern.Short_DateTime
                Return NewID("date_time")
        End Select
    End Function
    ''' <summary>
    ''' Returns a unique ID, depending on the pattern.
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function NewGUID(IDPattern_ As IDPattern, Optional prefx As String = "") As String
        Select Case IDPattern_
            Case IDPattern.Long_
                Return NewGUID(prefx)
            Case IDPattern.Long_DateTime
                Return NewGUID(prefx, True)
            Case IDPattern.Short_
                Return NewID("", prefx)
            Case IDPattern.Short_DateOnly
                Return NewID("acct")
            Case IDPattern.Short_DateTime
                Return NewID("date_time")
        End Select
    End Function
    ''' <summary>
    ''' Returns a unique ID, depending on the pattern.
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function NewGUID(Optional prefx As String = "", Optional suffx As Boolean = False) As String
        Dim return__ As String = ""
        If prefx.Length > 0 Then return__ = prefx & "-"
        If suffx = True Then return__ &= Now.ToShortDateString & "-" & Now.ToLongTimeString & "-"
        return__ = return__.Replace(" ", "")
        Dim r As String = return__.Replace("/", "-")
        Dim r_ As String = r.Replace(":", "-")
        Dim r3 As String = r_.Replace("AM", "-am")
        Dim return_ As String = r3.Replace("PM", "-pm")

        If return_.Length > 0 Then
            Return return_ & System.Guid.NewGuid().ToString
        Else
            Return System.Guid.NewGuid().ToString
        End If
    End Function
    Public Shared Function CalculateSince(date_ As Date, Optional interval_ As DateInterval = DateInterval.Day, Optional suffixed As Boolean = True) As String
        Dim d = DateDiff(interval_, date_, Date.Parse(Now))
        If suffixed Then
            Return ToPlural(d, "days") & " ago"
        Else
            Return d
        End If
    End Function
    Public Shared Function Acronym(string_ As String, Optional return_upper_case As Boolean = True, Optional separator As String = " ")
        If string_.Length < 1 Or separator.Length < 1 Then Return ""

        Dim s As Array = SplitTextInSplits(string_, separator, SideToReturn.AsArray)
        Dim r As String = ""
        For i = 0 To s.Length - 1
            r &= Mid(s(i), 1, 1)
        Next
        If return_upper_case Then
            Return r.ToUpper
        Else
            Return r
        End If
    End Function

    Public Shared Function ToPlural(count_ As Long, str_to_change As SingularWord, Optional rest_of_full_string As String = "", Optional prefixed As Boolean = True, Optional textCase As TextCase = TextCase.LowerCase, Optional prepend_with_approximately_if_count_is_of_type_double As Boolean = False) As String

        If prepend_with_approximately_if_count_is_of_type_double = True And count_ Mod 2 <> 0 Then
            Return "approximately " & ToPlural(count_, str_to_change.ToString, rest_of_full_string, prefixed, textCase)
        Else
            Return ToPlural(count_, str_to_change.ToString, rest_of_full_string, prefixed, textCase)
        End If
    End Function

    Private Shared Function ToPlural(count_ As Long, str_to_change As String, Optional rest_of_full_string As String = "", Optional prefixed As Boolean = True, Optional textCase As TextCase = TextCase.LowerCase) As String
        Dim c = 0
        Try
            If count_.ToString.Length > 0 Then c = count_
        Catch
        End Try

        Dim val_ As String = ""

        If Val(c) = 1 Then
            Select Case str_to_change.ToString.ToLower
                Case "message"
                    If prefixed Then
                        val_ = "a new message"
                    Else
                        val_ = "message"
                    End If
                Case "messages"
                    If prefixed Then
                        val_ = "a new message"
                    Else
                        val_ = "message"
                    End If
                Case "mark"
                    If prefixed Then
                        val_ = "1 mark "
                    Else
                        val_ = "mark"
                    End If
                Case "marks"
                    If prefixed Then
                        val_ = "1 mark "
                    Else
                        val_ = "mark"
                    End If
                Case "minute"
                    If prefixed Then
                        val_ = "1 minute "
                    Else
                        val_ = "minute"
                    End If
                Case "minutes"
                    If prefixed Then
                        val_ = "1 minute "
                    Else
                        val_ = "minute"
                    End If

                Case "day"
                    If prefixed Then
                        val_ = "1 day "
                    Else
                        val_ = "day"
                    End If
                Case "days"
                    If prefixed Then
                        val_ = "1 day "
                    Else
                        val_ = "day"
                    End If

                Case "student"
                    If prefixed Then
                        val_ = "1 student "
                    Else
                        val_ = "student"
                    End If
                Case "students"
                    If prefixed Then
                        val_ = "1 student "
                    Else
                        val_ = "student"
                    End If

                Case "month"
                    If prefixed Then
                        val_ = "1 month "
                    Else
                        val_ = "month"
                    End If
                Case "months"
                    If prefixed Then
                        val_ = "1 month "
                    Else
                        val_ = "month"
                    End If

                Case "year"
                    If prefixed Then
                        val_ = "1 year "
                    Else
                        val_ = "year"
                    End If
                Case "years"
                    If prefixed Then
                        val_ = "1 year "
                    Else
                        val_ = "year"
                    End If

                Case "hour"
                    If prefixed Then
                        val_ = "1 hour "
                    Else
                        val_ = "hour"
                    End If
                Case "hours"
                    If prefixed Then
                        val_ = "1 hour "
                    Else
                        val_ = "hour"
                    End If

                Case "second"
                    If prefixed Then
                        val_ = "1 second "
                    Else
                        val_ = "second"
                    End If
                Case "seconds"
                    If prefixed Then
                        val_ = "1 second "
                    Else
                        val_ = "second"
                    End If

                Case "account"
                    If prefixed Then
                        val_ = "1 account "
                    Else
                        val_ = "account"
                    End If
                Case "accounts"
                    If prefixed Then
                        val_ = "1 account "
                    Else
                        val_ = "account"
                    End If

                Case "male"
                    If prefixed Then
                        val_ = "1 male "
                    Else
                        val_ = "male"
                    End If
                Case "males"
                    If prefixed Then
                        val_ = "1 male "
                    Else
                        val_ = "male"
                    End If

                Case "female"
                    If prefixed Then
                        val_ = "1 female "
                    Else
                        val_ = "female"
                    End If
                Case "females"
                    If prefixed Then
                        val_ = "1 female "
                    Else
                        val_ = "female"
                    End If

                Case "user"
                    If prefixed Then
                        val_ = "1 user "
                    Else
                        val_ = "user"
                    End If
                Case "users"
                    If prefixed Then
                        val_ = "1 user "
                    Else
                        val_ = "user"
                    End If

                Case "file"
                    If prefixed Then
                        val_ = "1 file "
                    Else
                        val_ = "file"
                    End If
                Case "files"
                    If prefixed Then
                        val_ = "1 file "
                    Else
                        val_ = "file"
                    End If

                Case "record"
                    If prefixed Then
                        val_ = "1 record "
                    Else
                        val_ = "record"
                    End If
                Case "records"
                    If prefixed Then
                        val_ = "1 record "
                    Else
                        val_ = "record"
                    End If

                Case "question"
                    If prefixed Then
                        val_ = "1 question "
                    Else
                        val_ = "question"
                    End If
                Case "questions"
                    If prefixed Then
                        val_ = "1 question "
                    Else
                        val_ = "question"
                    End If

                Case "task"
                    If prefixed Then
                        val_ = "1 task "
                    Else
                        val_ = "task"
                    End If
                Case "tasks"
                    If prefixed Then
                        val_ = "1 task "
                    Else
                        val_ = "task"
                    End If

                Case "item"
                    If prefixed Then
                        val_ = "1 item "
                    Else
                        val_ = "item"
                    End If
                Case "items"
                    If prefixed Then
                        val_ = "1 item "
                    Else
                        val_ = "item"
                    End If

                Case "unit"
                    If prefixed Then
                        val_ = "1 unit "
                    Else
                        val_ = "unit"
                    End If
                Case "units"
                    If prefixed Then
                        val_ = "1 unit "
                    Else
                        val_ = "unit"
                    End If

                Case "room"
                    If prefixed Then
                        val_ = "1 room "
                    Else
                        val_ = "room"
                    End If
                Case "rooms"
                    If prefixed Then
                        val_ = "1 room "
                    Else
                        val_ = "room"
                    End If

                Case "client"
                    If prefixed Then
                        val_ = "1 client "
                    Else
                        val_ = "client"
                    End If
                Case "clients"
                    If prefixed Then
                        val_ = "1 client "
                    Else
                        val_ = "client"
                    End If

                Case "match"
                    If prefixed Then
                        val_ = "1 match "
                    Else
                        val_ = "match"
                    End If
                Case "matches"
                    If prefixed Then
                        val_ = "1 match "
                    Else
                        val_ = "match"
                    End If

                Case "invoice"
                    If prefixed Then
                        val_ = "1 invoice "
                    Else
                        val_ = "invoice"
                    End If
                Case "invoices"
                    If prefixed Then
                        val_ = "1 invoice "
                    Else
                        val_ = "invoice"
                    End If

                Case "sale"
                    If prefixed Then
                        val_ = "1 sale "
                    Else
                        val_ = "sale"
                    End If
                Case "sales"
                    If prefixed Then
                        val_ = "1 sale "
                    Else
                        val_ = "sale"
                    End If

                Case "receipt"
                    If prefixed Then
                        val_ = "1 receipt "
                    Else
                        val_ = "receipt"
                    End If
                Case "receipts"
                    If prefixed Then
                        val_ = "1 receipt "
                    Else
                        val_ = "receipt"
                    End If

                Case "product"
                    If prefixed Then
                        val_ = "1 product "
                    Else
                        val_ = "product"
                    End If
                Case "products"
                    If prefixed Then
                        val_ = "1 product "
                    Else
                        val_ = "product"
                    End If

            End Select
        Else
            Select Case str_to_change.ToString.ToLower
                Case "message"
                    If prefixed Then
                        val_ = c & " new messages"
                    Else
                        val_ = "messages"
                    End If
                Case "messages"
                    If prefixed Then
                        val_ = c & " new messages"
                    Else
                        val_ = "messages"
                    End If
                Case "mark"
                    If prefixed Then
                        val_ = c & " marks"
                    Else
                        val_ = "marks"
                    End If
                Case "marks"
                    If prefixed Then
                        val_ = c & " marks"
                    Else
                        val_ = "marks"
                    End If
                Case "minute"
                    If prefixed Then
                        val_ = c & " minutes"
                    Else
                        val_ = "minutes"
                    End If
                Case "minutes"
                    If prefixed Then
                        val_ = c & " minutes"
                    Else
                        val_ = "minutes"
                    End If

                Case "day"
                    If prefixed Then
                        val_ = c & " days"
                    Else
                        val_ = "days"
                    End If
                Case "days"
                    If prefixed Then
                        val_ = c & " days"
                    Else
                        val_ = "days"
                    End If

                Case "student"
                    If prefixed Then
                        val_ = c & " students"
                    Else
                        val_ = "students"
                    End If
                Case "students"
                    If prefixed Then
                        val_ = c & " students"
                    Else
                        val_ = "students"
                    End If

                Case "month"
                    If prefixed Then
                        val_ = c & " months"
                    Else
                        val_ = "months"
                    End If
                Case "months"
                    If prefixed Then
                        val_ = c & " months"
                    Else
                        val_ = "months"
                    End If
                Case "year"
                    If prefixed Then
                        val_ = c & " years"
                    Else
                        val_ = "years"
                    End If
                Case "years"
                    If prefixed Then
                        val_ = c & " years"
                    Else
                        val_ = "years"
                    End If
                Case "hour"
                    If prefixed Then
                        val_ = c & " hours"
                    Else
                        val_ = "hours"
                    End If
                Case "hours"
                    If prefixed Then
                        val_ = c & " hours"
                    Else
                        val_ = "hours"
                    End If
                Case "second"
                    If prefixed Then
                        val_ = c & " seconds"
                    Else
                        val_ = "seconds"
                    End If
                Case "seconds"
                    If prefixed Then
                        val_ = c & " seconds"
                    Else
                        val_ = "seconds"
                    End If
                Case "account"
                    If prefixed Then
                        val_ = c & " accounts"
                    Else
                        val_ = "accounts"
                    End If
                Case "accounts"
                    If prefixed Then
                        val_ = c & " accounts"
                    Else
                        val_ = "accounts"
                    End If

                Case "male"
                    If prefixed Then
                        val_ = c & " males"
                    Else
                        val_ = "males"
                    End If
                Case "males"
                    If prefixed Then
                        val_ = c & " males"
                    Else
                        val_ = "males"
                    End If
                Case "female"
                    If prefixed Then
                        val_ = c & " females"
                    Else
                        val_ = "females"
                    End If
                Case "females"
                    If prefixed Then
                        val_ = c & " females"
                    Else
                        val_ = "females"
                    End If
                Case "user"
                    If prefixed Then
                        val_ = c & " users"
                    Else
                        val_ = "users"
                    End If
                Case "users"
                    If prefixed Then
                        val_ = c & " users"
                    Else
                        val_ = "users"
                    End If
                Case "file"
                    If prefixed Then
                        val_ = c & " files"
                    Else
                        val_ = "files"
                    End If
                Case "files"
                    If prefixed Then
                        val_ = c & " files"
                    Else
                        val_ = "files"
                    End If
                Case "record"
                    If prefixed Then
                        val_ = c & " records"
                    Else
                        val_ = "records"
                    End If
                Case "records"
                    If prefixed Then
                        val_ = c & " records"
                    Else
                        val_ = "records"
                    End If

                Case "question"
                    If prefixed Then
                        val_ = c & " questions"
                    Else
                        val_ = "questions"
                    End If
                Case "questions"
                    If prefixed Then
                        val_ = c & " questions"
                    Else
                        val_ = "questions"
                    End If

                Case "task"
                    If prefixed Then
                        val_ = c & " tasks"
                    Else
                        val_ = "tasks"
                    End If
                Case "tasks"
                    If prefixed Then
                        val_ = c & " tasks"
                    Else
                        val_ = "tasks"
                    End If

                Case "item"
                    If prefixed Then
                        val_ = c & " items"
                    Else
                        val_ = "items"
                    End If
                Case "items"
                    If prefixed Then
                        val_ = c & " items"
                    Else
                        val_ = "items"
                    End If

                Case "unit"
                    If prefixed Then
                        val_ = c & " units"
                    Else
                        val_ = "units"
                    End If
                Case "units"
                    If prefixed Then
                        val_ = c & " units"
                    Else
                        val_ = "units"
                    End If

                Case "room"
                    If prefixed Then
                        val_ = c & " rooms"
                    Else
                        val_ = "rooms"
                    End If
                Case "rooms"
                    If prefixed Then
                        val_ = c & " rooms"
                    Else
                        val_ = "rooms"
                    End If

                Case "client"
                    If prefixed Then
                        val_ = c & " clients"
                    Else
                        val_ = "clients"
                    End If
                Case "clients"
                    If prefixed Then
                        val_ = c & " clients"
                    Else
                        val_ = "clients"
                    End If

                Case "match"
                    If prefixed Then
                        val_ = c & " matches"
                    Else
                        val_ = "matches"
                    End If
                Case "matches"
                    If prefixed Then
                        val_ = c & " matches"
                    Else
                        val_ = "matches"
                    End If

                Case "invoice"
                    If prefixed Then
                        val_ = c & " invoices"
                    Else
                        val_ = "invoices"
                    End If
                Case "invoices"
                    If prefixed Then
                        val_ = c & " invoices"
                    Else
                        val_ = "invoices"
                    End If

                Case "sale"
                    If prefixed Then
                        val_ = c & " sales"
                    Else
                        val_ = "sales"
                    End If
                Case "sales"
                    If prefixed Then
                        val_ = c & " sales"
                    Else
                        val_ = "sales"
                    End If

                Case "receipt"
                    If prefixed Then
                        val_ = c & " receipts"
                    Else
                        val_ = "receipts"
                    End If
                Case "receipts"
                    If prefixed Then
                        val_ = c & " receipts"
                    Else
                        val_ = "receipts"
                    End If

                Case "product"
                    If prefixed Then
                        val_ = c & " products"
                    Else
                        val_ = "products"
                    End If
                Case "products"
                    If prefixed Then
                        val_ = c & " products"
                    Else
                        val_ = "products"
                    End If

            End Select
        End If

        If rest_of_full_string.Length > 0 Then
            val_ &= " " & rest_of_full_string
        End If

        Return TransformText(val_, textCase)
    End Function

    ''' <summary>
    ''' Determines if val_ is within 0 and 100.
    ''' </summary>
    ''' <param name="val_"></param>
    ''' <returns></returns>
    Public Shared Function InPercent(val_ As String) As Boolean
        val_ = val_.Trim
        If Val(val_) < 0 Or Val(val_) > 100 Then ' Or Val(val_) = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Shared Function IsEven(val As Long) As Boolean
        IsEven = (val Mod 2 = 0)
    End Function
    Public Shared Function ToBString(val) As String
        Select Case val.ToString
            Case True
                Return "Yes"
            Case False
                Return "No"
        End Select
    End Function
    ''' <summary>
    ''' Checks if PC is plugged in.
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function MachineIsOnAC() As Boolean
        Dim myManagedPower As New ManagedPower()
        Return myManagedPower.ToString().ToLower = "ac"
    End Function
    ''' <summary>
    ''' Checks if PC is not plugged in (i.e. running on battery).
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function MachineIsOnBattery() As Boolean
        Dim myManagedPower As New ManagedPower()
        Return myManagedPower.ToString().ToLower <> "ac"
    End Function
    Public Shared Function ReplaceText(body_of_text As String, word_to_remove As String, Optional what_to_replace_with As String = "")
        Return body_of_text.Replace(word_to_remove, what_to_replace_with)
    End Function
    Public Shared Function ReplaceText(body_of_text As String, list_of_words_to_remove As Array, Optional what_to_replace_with As String = "")
        Dim current_result As String
        Dim result As String = body_of_text
        For i = 0 To list_of_words_to_remove.Length - 1
            current_result = result.Replace(list_of_words_to_remove(i), what_to_replace_with)
            result = current_result
        Next
        Return result
    End Function

    Public Shared Function ReplaceText(body_of_text As String, list_of_words_to_remove As List(Of String), Optional what_to_replace_with As String = "")
        Return ReplaceText(body_of_text, list_of_words_to_remove.ToArray, what_to_replace_with)
    End Function

    ''' <summary>
    ''' Creates, adds to, or overwrites the content of file (format is text).
    ''' </summary>
    ''' <param name="file_">The path to the file.</param>
    ''' <param name="txt_">Intended string content of the file.</param>
    ''' <param name="append_">Should it add to the content of the file (if it has) or overwrite everything?</param>
    ''' <param name="dont_trim">Should trailing spaces be ignored?</param>
    Public Shared Function WriteText(file_ As String, txt_ As String, Optional append_ As Boolean = False, Optional dont_trim As Boolean = False)
        'If file_.Length < 1 Then Return

        'Dim t As String = txt_
        'If dont_trim = False Then t = txt_.Trim
        'Try
        '	My.Computer.FileSystem.WriteAllText(file_, t, append_)
        'Catch ex As Exception
        'End Try



        Dim [FileVar] As String = file_
        If append_ = True Then
            FileOpen(1, [FileVar], OpenMode.Append)
        Else
            FileOpen(1, [FileVar], OpenMode.Output)
        End If
        If dont_trim Then
            PrintLine(1, txt_)
        Else
            PrintLine(1, txt_.Trim)
        End If
        FileClose(1)

    End Function

    ''' <summary>
    ''' Retrieves the content of a file (format is text).
    ''' </summary>
    ''' <param name="file_">The path to the file.</param>
    ''' <param name="dont_trim">Should trailing spaces be ignored?</param>
    ''' <returns>The (text) content of the file.</returns>
    Public Shared Function ReadText(file_ As String, Optional dont_trim As Boolean = True) As String
        'If file_ Is Nothing Then Exit Function

        'If CType(file_, String).Length < 1 Then Exit Function

        If file_.Length < 1 Then Return ""

        Dim r As String = My.Computer.FileSystem.ReadAllText(file_).Trim
        If dont_trim = True Then
            r = My.Computer.FileSystem.ReadAllText(file_)
        End If
        Return r




        'Dim docName As String = Path.GetFileName(file_)
        'Dim docPath As String = Path.GetDirectoryName(file_)
        'Dim stream As New FileStream(file_, FileMode.Open)
        'Dim reader As New StreamReader(stream)
        'Try
        '	If dont_trim = True Then
        '		Return reader.ReadToEnd()
        '	Else
        '		Return reader.ReadToEnd().Trim
        '	End If
        'Catch
        'Finally
        '	reader.Dispose()
        '	stream.Dispose()
        'End Try
        'Return True


2:

    End Function

    Public Shared Sub Delete(file_ As String, Optional recycle_ As Boolean = True, Optional showUI_ As Boolean = False)
        Try
            My.Computer.FileSystem.DeleteFile(file_, showUI_, recycle_)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub Delete(files_ As Array, Optional recycle_ As Boolean = True, Optional showUI_ As Boolean = False)
        If files_.Length > 0 Then
            For i = 0 To files_.Length - 1
                Delete(files_(i), recycle_, showUI_)
            Next
        End If
    End Sub

    ''' <summary>
    ''' Set permissions on file
    ''' </summary>
    ''' <param name="file_"></param>
    ''' <param name="perm_"></param>
    ''' <param name="remove_existing"></param>
    Public Sub PermissionForFile(file_ As String, Optional perm_ As FileSystemRights = FileSystemRights.FullControl, Optional remove_existing As Boolean = False)
        If file_.Length < 1 Then Exit Sub
        If perm_.ToString.Length < 1 Then perm_ = FileSystemRights.FullControl

        Dim user_ As String = Environment.UserName

        Dim FilePath As String = file_
        Dim UserAccount As String = user_
        Dim FileInfo As IO.FileInfo = New IO.FileInfo(FilePath)
        Dim FileAcl As New FileSecurity
        FileAcl.AddAccessRule(New FileSystemAccessRule(UserAccount, perm_, AccessControlType.Deny))
    End Sub
    ''' <summary>
    ''' Set permissions on folder
    ''' </summary>
    ''' <param name="folder_">Folder to set permission on</param>
    ''' <param name="perm_">Type of permission</param>
    ''' <param name="remove_existing">Should existing permissions for this user be overwritten if found?</param>
    Public Sub PermitForFolder(folder_ As String, Optional perm_ As FileSystemRights = FileSystemRights.FullControl, Optional remove_existing As Boolean = False)
        If folder_.Length < 1 Then Exit Sub
        If perm_.ToString.Length < 1 Then perm_ = FileSystemRights.FullControl

        Dim user_ As String = Environment.UserName

        Dim FolderPath As String = folder_
        Dim UserAccount As String = user_

        Dim FolderInfo As IO.DirectoryInfo = New IO.DirectoryInfo(FolderPath)
        Dim FolderAcl As New DirectorySecurity
        FolderAcl.AddAccessRule(New FileSystemAccessRule(UserAccount, perm_, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
        If remove_existing = True Then FolderAcl.SetAccessRuleProtection(True, False)
        FolderInfo.SetAccessControl(FolderAcl)
    End Sub

    ''' <summary>
    ''' file_ can now start/run with Windows logon. Same as ToStartup.
    ''' </summary>
    ''' <param name="file_"></param>
    ''' <param name="key_"></param>

    Private Shared Sub ToMachineStartup(file_ As String, key_ As String)
        If file_.Length < 1 Or key_.Length < 1 Then Exit Sub
        Try
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", key_, file_)
        Catch x As Exception
        End Try
    End Sub
    ''' <summary>
    ''' file_ can now start/run with Windows logon. Same as ToMachineStartup.
    ''' </summary>
    ''' <param name="file_"></param>
    ''' <param name="key_"></param>
    Public Shared Sub ToStartup(file_ As String, key_ As String)
        ToMachineStartup(file_, key_)
    End Sub

    Public Sub RemoveFromStartup(file_ As String, key_ As String)
        If file_.Length < 1 Or key_.Length < 1 Then Exit Sub

        Using key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software")
            key.DeleteSubKey("Software\Microsoft\Windows\CurrentVersion\Run\" & key_)
        End Using

        Dim exists As Boolean = False
        Try
            If My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run\" & key_) IsNot Nothing Then
            End If
        Catch
        Finally
        End Try

    End Sub
    Public Shared Sub RegistryRemoveKey(section As RegistryKey, subKey_ As String)
        Try
            section.DeleteSubKey(subKey_)
        Catch ex As Exception
        End Try

    End Sub

    Public Sub SetAttribute(file_OR_directory As String, Optional show_ As Boolean = False, Optional remove_ As Boolean = True)
        If show_ = False And remove_ = False Then
            Try
                SetAttr(file_OR_directory, FileAttribute.Hidden)
                Exit Sub
            Catch ex As Exception
            End Try
        End If

        If show_ = True Then
            Try
                SetAttr(file_OR_directory, FileAttribute.Normal)
                Exit Sub
            Catch ex As Exception
            End Try
        Else
        End If

        If remove_ = True Then
            Try
                SetAttr(file_OR_directory, FileAttribute.Hidden + FileAttribute.System)
                Exit Sub
            Catch ex As Exception
            End Try
        Else
        End If

    End Sub

    Public Shared Sub CreateShortcut(target_ As String, folder_ As String, filename_without_extension As String, Optional icon_file As String = Nothing)
        Dim wsh As Object = CreateObject("WScript.Shell")

        wsh = CreateObject("WScript.Shell")

        Dim MyShortcut, DesktopPath

        ' Read desktop path using WshSpecialFolders object

        DesktopPath = folder_

        MyShortcut = wsh.CreateShortcut(DesktopPath & "\" & filename_without_extension & ".lnk")

        ' Set shortcut object properties and save it

        MyShortcut.TargetPath = wsh.ExpandEnvironmentStrings(target_)

        MyShortcut.WorkingDirectory = wsh.ExpandEnvironmentStrings(folder_)

        MyShortcut.WindowStyle = 4

        If icon_file IsNot Nothing Then
            If icon_file.Length > 1 Then
                MyShortcut.IconLocation = wsh.ExpandEnvironmentStrings(icon_file)
            End If
        End If

        'Save the shortcut

        MyShortcut.Save()
    End Sub

    Public Shared Function AppIsOn(file_path As String) As Boolean
        Dim p() As Process = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(file_path.Trim))
        AppIsOn = p.Count > 0
    End Function
    Public Shared Function AppNotOn(file_path As String) As Boolean
        Dim p() As Process = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(file_path.Trim))
        AppNotOn = p.Count < 1
    End Function

    Public Shared Sub KillProcess(process_name_without_extension As String)
        For Each proc As Process In Process.GetProcessesByName(process_name_without_extension)
            proc.Kill()
        Next
    End Sub
    Public Shared Sub TaskManager()
        StartFile("taskmgr")
    End Sub

    <DllImport("User32.dll")>
    Public Shared Function ShowWindow(ByVal hwnd As IntPtr, ByVal cmd As Integer) As Boolean

    End Function

    Public Shared Function ShowApp(process_name As String, state_ As AppWindowState) As Boolean
        If process_name.Length < 1 Then Return False

        process_name = Path.GetFileNameWithoutExtension(process_name)

        Try
            Dim p() As Process = Process.GetProcessesByName(process_name)
            For Each pr As Process In p
                ShowWindow(pr.MainWindowHandle, state_)
            Next
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Shared Function ShowDesktop(Optional show_ As Boolean = False) As Boolean
        Dim show_or_hide As Integer = 6
        If show_ = True Then show_or_hide = 9

        Try
            For Each p As Process In Process.GetProcesses
                ShowWindow(p.MainWindowHandle, show_or_hide)
            Next
            Return True
        Catch
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Opens a file. If checkFirst is True, then the program file won't run if it is already running; this is ignored if it's not a program file.
    ''' </summary>
    ''' <param name="file_">The path to the file.</param>
    ''' <param name="style_">ProcessWindowStyle (normal, hidden etc)</param>
    ''' <param name="style__">Ignore ProcessWindowStyle, completely hide it.</param>
    ''' <param name="checkFirst">Check if an instance of the file is already running (in Tasks).</param>
    ''' <returns>True if file exists and is opened, false if not.</returns>
    Public Shared Function StartFile(file_ As String, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False, Optional checkFirst As Boolean = False) As Boolean
        If checkFirst = True Then
            Try
                If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
                    Exit Function
                End If
            Catch
            End Try
        End If

        Dim s As New ProcessStartInfo
        If style__ = True Then
            s.WindowStyle = ProcessWindowStyle.Hidden
        Else
            s.WindowStyle = style_
        End If
        s.FileName = file_
        Try
            Process.Start(s)
            Return True
        Catch ex As Exception
            Return False
        End Try

2:
    End Function
    ''' <summary>
    ''' Opens a file with argument. If checkFirst is True, then the program file won't run if it is already running; this is ignored if it's not a program file. Same as StartFileWithArgument.
    ''' </summary>
    ''' <param name="file_"></param>
    ''' <param name="arg_"></param>
    ''' <param name="checkFirst"></param>
    ''' <returns></returns>
    Public Shared Function StartFile(file_ As String, arg_ As String, Optional checkFirst As Boolean = False) As Boolean
        If checkFirst = True Then
            Try
                If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
                    Exit Function
                End If
            Catch
            End Try
        End If

        Try
            Process.Start(file_, arg_)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function StartFileWithArgument(file_ As String, Optional arg_ As String = Nothing, Optional checkFirst As Boolean = False) As Boolean
        If checkFirst = True Then
            Try
                If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
                    Exit Function
                End If
            Catch
            End Try
        End If

        Try
            Process.Start(file_, arg_)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function


    Public Shared Sub LogOff()
        Try
            Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-l")
        Catch
        End Try
    End Sub
    Public Shared Sub Restart()
        Try
            Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-r -f -t 0")
        Catch
        End Try
    End Sub
    Public Shared Sub Shutdown()
        Try
            Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-s -f -t 0")
        Catch
        End Try
    End Sub

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Friend Shared Sub LockWorkStation()

    End Sub
    Public Shared Sub LockPC()
        LockWorkStation()
    End Sub

    Public Shared Sub ExitApp(Optional process_name_without_extension As String = "")
        'taskkill /f /im "explorer.exe"

        Try
            Select Case process_name_without_extension.Length
                Case < 1
                    Environment.Exit(0)
                Case Else
                    If AppIsOn(process_name_without_extension) Then KillProcess(process_name_without_extension)
            End Select
        Catch
        End Try
    End Sub

    Public Shared Sub CopyFiles(src As String, tgt As String, Optional dialogs As FileIO.UIOption = FileIO.UIOption.AllDialogs)
        If src.Trim.Length < 1 Or tgt.Trim.Length < 1 Then Return
        My.Computer.FileSystem.CopyDirectory(src, tgt, dialogs, FileIO.UICancelOption.DoNothing)
    End Sub

    Public Shared Shadows Function RenameFile(src_file_path, tgt_file_name_plus_extension) As Boolean
        Try
            My.Computer.FileSystem.RenameFile(src_file_path, tgt_file_name_plus_extension)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    ''' <summary>
    ''' Checks if a file exists.
    ''' </summary>
    ''' <param name="file_"></param>
    ''' <returns></returns>
    Public Shared Function Exists(file_ As String) As Boolean
        If file_.Length < 1 Then
            Return False
            Exit Function
        End If
        Exists = My.Computer.FileSystem.FileExists(file_) Or My.Computer.FileSystem.DirectoryExists(file_)
    End Function

    Private Shared Function TransformWord(word As Object, casing As TextCase)
        If CType(word, String).Length < 1 Then Return ""
        Dim s As String = CStr(word)
        Dim r = ""
        Select Case casing
            Case TextCase.Capitalize
                r = Mid(s, 1, 1).ToUpper & Mid(s, 2).ToLower
            Case TextCase.LowerCase
                r = s.ToLower
            Case TextCase.UpperCase
                r = s.ToUpper
            Case TextCase.None
                r = s
        End Select
        Return r
    End Function
    Private Shared Function TransformSingleLineText(text As Object, Optional casing As TextCase = TextCase.Capitalize, Optional separator_ As String = " ")
        If CType(text, String).Trim.Length < 1 Then Return ""

        Dim d As List(Of String) = SplitTextInSplits(CStr(text), separator_, SideToReturn.AsListOfString)
        Dim o As New List(Of String)
        With d
            For i = 0 To .Count - 1
                o.Add(TransformWord(d(i), casing))
            Next
        End With
        Dim r = ""
        With o
            For i = 0 To .Count - 1
                r &= o(i)
                If i <> .Count - 1 Then r &= separator_
            Next
        End With
        Return r.Trim
    End Function

    Private Shared Function TransformMultiLineText(text As Object, Optional casing As TextCase = TextCase.Capitalize, Optional separator_ As String = " ")
        Dim s As String = CStr(text).Trim
        Dim o As List(Of String) = SplitTextInSplits(s, vbCrLf, SideToReturn.AsListOfString)
        Dim f As New List(Of String)
        With o
            For i = 0 To .Count - 1
                f.Add(TransformSingleLineText(o(i).Trim, casing, separator_))
            Next
        End With
        Return ListToString(f)
    End Function
    Public Shared Function TransformText(text As Object, Optional casing As TextCase = TextCase.Capitalize, Optional separator_ As String = " ")
        Dim s As String = CStr(text)
        If s.Length < 1 Then Return ""
        If s.Contains(vbCrLf) Then
            Return TransformMultiLineText(text, casing, separator_)
        Else
            Return TransformSingleLineText(text, casing, separator_)
        End If
    End Function

    ''' <summary>
    ''' Checks if a string is valid email address.
    ''' </summary>
    ''' <param name="email_">String to check</param>
    ''' <returns>True or False</returns>
    Public Shared Function IsEmail(email_ As String) As Boolean
        Dim isValid As Boolean = True
        If Not Regex.IsMatch(email_,
            "^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$") Then
            isValid = False
        End If
        Return isValid
    End Function
    Public Shared Function firstWord(text As String) As String
        If text.Length < 1 Then Return ""

        Return SplitTextInTwo(text.Trim, " ", SideToReturn.Left)
    End Function

    Public Shared Function otherWords(text As String) As String
        If text.Length < 1 Then Return ""

        Return SplitTextInTwo(text.Trim, " ", SideToReturn.Right)
    End Function

    Public Shared Function lastThreeLetters(text As String) As Array
        Return {text.Chars(text.Length - 3), text.Chars(text.Length - 2), text.Chars(text.Length - 1)}
    End Function

    Public Shared Function IsVowel(text As String) As Boolean
        If text.Length < 1 Then Return False
        If text.ToLower = "a" Or
                 text.ToLower = "e" Or
                 text.ToLower = "i" Or
                 text.ToLower = "o" Or
                 text.ToLower = "u" Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Function IsConsonant(text As String) As Boolean
        If text.Length < 1 Then Return False
        If text.ToLower = "b" Or
                 text.ToLower = "c" Or
                 text.ToLower = "d" Or
                 text.ToLower = "f" Or
                 text.ToLower = "g" Or
                 text.ToLower = "h" Or
                 text.ToLower = "j" Or
                 text.ToLower = "k" Or
                 text.ToLower = "l" Or
                 text.ToLower = "m" Or
                 text.ToLower = "n" Or
                 text.ToLower = "p" Or
                 text.ToLower = "q" Or
                 text.ToLower = "r" Or
                 text.ToLower = "s" Or
                 text.ToLower = "t" Or
                 text.ToLower = "v" Or
                 text.ToLower = "w" Or
                 text.ToLower = "x" Or
                 text.ToLower = "y" Or
                 text.ToLower = "z" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function IsAlphabet(text As String) As Boolean
        If IsConsonant(text) Or IsVowel(text) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Turns statement to continuous tense. Works for ~80% scenarios.
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="suffx"></param>
    ''' <returns></returns>
    Public Shared Function ToContinuous(text As String, Optional suffx As String = "") As String
        If text.Trim.Length < 1 Then Return ""
        Dim lastThree As Array = lastThreeLetters(text)
        If IsAlphabet(lastThree(0)) = False Or
                IsAlphabet(lastThree(1)) = False Or
                IsAlphabet(lastThree(2)) = False Then
            Return ""
        End If

        Dim a = lastThree(0).ToString.ToLower, b = lastThree(1).ToString.ToLower, c = lastThree(2).ToString.ToLower
        Dim prefx = ""

        If a = "i" And b = "n" And c = "g" Then
            Return text
        End If

        If IsConsonant(a) And IsVowel(b) And IsConsonant(c) Then
            prefx = text & Mid(text.Trim, text.Length, 1).Trim & "ing"
        ElseIf b = "i" And c = "e" Then
            prefx = Mid(text.Trim, 1, text.Length - 2).Trim & "ying"
        ElseIf IsVowel(a) And IsConsonant(b) And c = "e" Then
            prefx = Mid(text.Trim, 1, text.Length - 1).Trim & "ing"
        Else
            prefx = text.Trim & "ing"
        End If

        Return RTrim(prefx) & " " & LTrim(suffx)
    End Function
    Public Shared Function IsPhraseOrSentence(text As String) As Boolean
        If text.Trim.Length < 1 Then Return False
        If firstWord(text).Length > 0 And otherWords(text).Length > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Function FullNameFromNames(FirstName As String, LastName As String, Optional TitleOfCourtesy As String = "", Optional MiddleName As String = "", Optional last_name_first As Boolean = False, Optional include_title_of_courtesy As Boolean = True, Optional separate_first_or_last_name_with_comma As Boolean = False, Optional include_middle_name As Boolean = False) As String
        Dim r As String = ""
        If last_name_first Then
            If include_middle_name = True And MiddleName.Trim.Length > 0 Then
                r = LastName.Trim & " " & FirstName.Trim & " " & MiddleName.Trim
                If separate_first_or_last_name_with_comma Then
                    r = LastName.Trim & ", " & FirstName.Trim & " " & MiddleName.Trim
                End If
            Else
                r = LastName.Trim & " " & FirstName.Trim
                If separate_first_or_last_name_with_comma Then
                    r = LastName.Trim & ", " & FirstName.Trim
                End If
            End If
        Else
            r = FirstName.Trim & " " & LastName.Trim
            If separate_first_or_last_name_with_comma Then
                r = FirstName.Trim & ", " & LastName.Trim
            End If
            If include_middle_name = True And MiddleName.Trim.Length > 0 Then
                r = FirstName.Trim & " " & MiddleName & " " & LastName.Trim
                If separate_first_or_last_name_with_comma Then
                    r = FirstName.Trim & ", " & MiddleName & " " & LastName.Trim
                End If
            End If
        End If
        If include_title_of_courtesy And TitleOfCourtesy.Trim.Length > 0 Then
            r = TitleOfCourtesy & " " & r
        End If
        Return r
    End Function

    Public Shared Function ContainsAlphabet(str) As Boolean
        Dim r As Boolean = False
        If CType(str, String).ToLower.Contains("a") Or
            CType(str, String).ToLower.Contains("b") Or
            CType(str, String).ToLower.Contains("c") Or
            CType(str, String).ToLower.Contains("d") Or
            CType(str, String).ToLower.Contains("e") Or
            CType(str, String).ToLower.Contains("f") Or
            CType(str, String).ToLower.Contains("g") Or
            CType(str, String).ToLower.Contains("h") Or
            CType(str, String).ToLower.Contains("i") Or
            CType(str, String).ToLower.Contains("j") Or
            CType(str, String).ToLower.Contains("k") Or
            CType(str, String).ToLower.Contains("l") Or
            CType(str, String).ToLower.Contains("m") Or
            CType(str, String).ToLower.Contains("n") Or
            CType(str, String).ToLower.Contains("o") Or
            CType(str, String).ToLower.Contains("p") Or
            CType(str, String).ToLower.Contains("q") Or
            CType(str, String).ToLower.Contains("r") Or
            CType(str, String).ToLower.Contains("s") Or
            CType(str, String).ToLower.Contains("t") Or
            CType(str, String).ToLower.Contains("u") Or
            CType(str, String).ToLower.Contains("v") Or
            CType(str, String).ToLower.Contains("w") Or
            CType(str, String).ToLower.Contains("x") Or
            CType(str, String).ToLower.Contains("y") Or
            CType(str, String).ToLower.Contains("z") Or
            CType(str, String).ToLower.Contains("-") Then
            r = True
        End If

        If str.ToString.Length < 1 Then r = False

        Return r
    End Function
    ''' <summary>
    ''' Gets all enum values and adds them to List(of String) which it returns. Each Enum value must be equal to a number, e.g. ... Value1=1  Value2 = 2
    ''' </summary>
    ''' <param name="instance_of_enum"></param>
    ''' <param name="values_instead"></param>
    ''' <returns></returns>
    Public Shared Function GetEnum(instance_of_enum As Object, Optional values_instead As Boolean = False) As List(Of String)
        Dim l As New List(Of String)

        If values_instead Then
            For Each i As Integer In [Enum].GetValues(instance_of_enum.GetType)
                l.Add(i)
            Next
            Return l
        End If

        For Each i In instance_of_enum.GetType().GetEnumNames()
            l.Add(i)
        Next
        Return l

    End Function

    Public Shared Function Percentage(score_, max_score)
        Return (score_ / max_score) * 100
    End Function

    Public Shared Function DictionaryKeys(dict As Dictionary(Of String, Object), Optional side_to_return As SideToReturn = SideToReturn.AsListOfString)
        Dim d As Dictionary(Of String, Object) = dict
        Dim l As New List(Of String)
        With d
            For i = 0 To .Count - 1
                l.Add(d.Keys(i))
            Next
        End With
        If side_to_return = SideToReturn.AsListOfString Then
            Return l
        ElseIf side_to_return = SideToReturn.AsArray Then
            Return l.ToArray
        End If
    End Function
    Public Shared Function DictionaryValues(dict As Dictionary(Of String, Object), Optional side_to_return As SideToReturn = SideToReturn.AsListOfString)
        Dim d As Dictionary(Of String, Object) = dict
        Dim l As New List(Of String)
        With d
            For i = 0 To .Count - 1
                l.Add(d.Values(i))
            Next
        End With
        If side_to_return = SideToReturn.AsListOfString Then
            Return l
        ElseIf side_to_return = SideToReturn.AsArray Then
            Return l.ToArray
        End If
    End Function

    Public Shared Function DictionaryToString(dict As Dictionary(Of String, Object), Optional side_to_return As SideToReturn = SideToReturn.Right)
        Dim d As Dictionary(Of String, Object) = dict
        Dim l As New List(Of String)
        Dim r As New List(Of Object)
        With d
            For i = 0 To .Count - 1
                If side_to_return = SideToReturn.AsArray Or side_to_return = SideToReturn.Left Then
                    l.Add(d.Keys(i))
                End If
                If side_to_return = SideToReturn.AsArray Or side_to_return = SideToReturn.Right Then
                    r.Add(d.Values(i))
                End If
            Next
        End With
        If side_to_return = SideToReturn.Left Then
            Return ListToString(l)
        ElseIf side_to_return = SideToReturn.Right Then
            Return ListToString(r)
        Else 'side_to_return = SideToReturn.AsArray
            Return {l, r}
        End If
    End Function

    Public Shared Function ListToString(list_or_array As Object, Optional delimiter As String = vbCrLf, Optional format_output As Boolean = False) As String
        Dim delimeter As String = delimiter
        Dim l 'As List(Of String)
        If TypeOf list_or_array Is Array Then
            Return ArrayToString(list_or_array)
        ElseIf TypeOf list_or_array Is List(Of Object) Then
            l = CType(list_or_array, List(Of Object))
        Else
            l = CType(list_or_array, List(Of String))
        End If
        Dim r As String = ""
        With l
            For i As Integer = 0 To .Count - 1
                r &= l(i) & delimeter
            Next
        End With
        If format_output Then
            Return PrepareForIO(r)
        Else
            Return r
        End If
    End Function
    Private Shared Function CustomMarkup(str_ As String) As String
        Return str_
    End Function
    Public Shared Function PrepareForIO(str_ As String, Optional output_ As FormatFor = FormatFor.Web) As String
        If output_ = FormatFor.Custom Then Return CustomMarkup(str_)

        Dim trimmed_, CR_less, CRLFless, TABless As String
        trimmed_ = str_.Trim
        str_ = trimmed_

        If output_ = FormatFor.Web Then
            CRLFless = str_.Replace(vbCrLf, "<br />")
            str_ = CRLFless
        ElseIf output_ = FormatFor.JavaScript Then
            CRLFless = str_.Replace(vbCrLf, "\n")
            str_ = CRLFless
        End If

        If output_ = FormatFor.Web Then
            CR_less = str_.Replace(vbCrLf, "<br />")
            str_ = CR_less
        ElseIf output_ = FormatFor.JavaScript Then
            CR_less = str_.Replace(vbCrLf, "\n")
            str_ = CR_less
        End If

        If output_ = FormatFor.Web Then
            TABless = str_.Replace(vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
            str_ = TABless
        ElseIf output_ = FormatFor.JavaScript Then
            TABless = str_.Replace(vbTab, "\t")
            str_ = TABless
        End If

        Return str_
    End Function

    Public Shared Function ArrayToString(array As Array, Optional delimeter As String = vbCrLf, Optional format_output As Boolean = False) As String
        Dim r As String = ""
        With array
            For i As Integer = 0 To .Length - 1
                r &= array(i) & delimeter
            Next
        End With
        If format_output Then
            Return PrepareForIO(r)
        Else
            Return r
        End If

    End Function

    Public Shared Function ArrayToList(array_ As Array, list_is As ListIsOf) As Object
        Dim a As Array = array_
        Dim l_string As New List(Of String)

        For Each i In a
            If list_is = ListIsOf.String_ Then
                l_string.Add(i)
            End If
        Next

        Return l_string
    End Function

    Public Shared Function StringToList(delimited_string As String, Optional delimeter As String = vbCrLf, Optional return_ As SideToReturn = SideToReturn.AsListOfString) As Object
        Return SplitTextInSplits(delimited_string, delimeter, return_)
    End Function

#Region "DropText-Boolean"


    Public Shared ReadOnly Property BooleanDropTextTrue As List(Of String) =
        {"yes, include this item", "always", "yes", "on", "one", "true", "if possible", "succeeded"}.ToList()
    Public Shared ReadOnly Property BooleanDropTextFalse As List(Of String) =
        {"no, remove this item", "never", "no", "off", "zero", "false", "not at all", "failed"}.ToList()
    Public Shared cart_inclusion_list As String() = {"Yes, Include This Item", "No, Remove This Item"}


    ''' <summary>
    ''' Converts boolean values to user-friendly pattern e.g. yes, never.
    ''' To convert to boolean, use DropTextToBoolean(str_ As String).
    ''' </summary>
    ''' <param name="boolean_val">True or False</param>
    ''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
    ''' <returns>Always/Never (default), Yes/No, On/Off, 1/0, True/Talse, If possible/Not at all</returns>
    Private Shared Function BooleanToDropText(boolean_val As Boolean, Optional pattern_ As String = "Yes/No") As String

        Select Case Convert.ToBoolean(boolean_val)
            Case True
                If pattern_.ToLower = "" Then
                    Return "Always"
                End If
                If pattern_.ToLower = "always/never" Then
                    Return "Always"
                End If
                If pattern_.ToLower = "a/n" Then
                    Return "Always"
                End If
                If pattern_.ToLower = "a" Then
                    Return "Always"
                End If
                If pattern_.ToLower = "yes/no" Then
                    Return "Yes"
                End If
                If pattern_.ToLower = "y/n" Then
                    Return "Yes"
                End If
                If pattern_.ToLower = "y" Then
                    Return "Yes"
                End If
                If pattern_.ToLower = "on/off" Then
                    Return "On"
                End If
                If pattern_.ToLower = "o/f" Then
                    Return "On"
                End If
                If pattern_.ToLower = "o" Then
                    Return "On"
                End If
                If pattern_.ToLower = "1/0" Then
                    Return "1"
                End If
                If pattern_.ToLower = "true/false" Then
                    Return "True"
                End If
                If pattern_.ToLower = "t/f" Then
                    Return "True"
                End If
                If pattern_.ToLower = "t" Then
                    Return "True"
                End If
                If pattern_.ToLower = "if/not" Then
                    Return "If possible"
                End If
                If pattern_.ToLower = "i/n" Then
                    Return "If possible"
                End If
                If pattern_.ToLower = "i" Then
                    Return "If possible"
                End If

            Case False
                If pattern_.ToLower = "" Then
                    Return "Never"
                End If
                If pattern_.ToLower = "always/never" Then
                    Return "Never"
                End If
                If pattern_.ToLower = "a/n" Then
                    Return "Never"
                End If
                If pattern_.ToLower = "a" Then
                    Return "Never"
                End If
                If pattern_.ToLower = "yes/no" Then
                    Return "No"
                End If
                If pattern_.ToLower = "y/n" Then
                    Return "No"
                End If
                If pattern_.ToLower = "y" Then
                    Return "No"
                End If
                If pattern_.ToLower = "on/off" Then
                    Return "Off"
                End If
                If pattern_.ToLower = "o/f" Then
                    Return "Off"
                End If
                If pattern_.ToLower = "o" Then
                    Return "Off"
                End If
                If pattern_.ToLower = "1/0" Then
                    Return "0"
                End If
                If pattern_.ToLower = "true/false" Then
                    Return "False"
                End If
                If pattern_.ToLower = "t/f" Then
                    Return "False"
                End If
                If pattern_.ToLower = "t" Then
                    Return "False"
                End If
                If pattern_.ToLower = "if/not" Then
                    Return "Not at all"
                End If
                If pattern_.ToLower = "i/n" Then
                    Return "Not at all"
                End If
                If pattern_.ToLower = "i" Then
                    Return "Not at all"
                End If
        End Select
    End Function

    Public Shared Function BooleanToDropText(boolean_val As Boolean, Optional pattern_ As DropTextPattern = DropTextPattern.YesNo) As String

        Select Case Convert.ToBoolean(boolean_val)
            Case True
                If pattern_.ToString.ToLower = DropTextPattern.AlwaysNever.ToString.ToLower Then
                    Return "Always"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.YesNo.ToString.ToLower Then
                    Return "Yes"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.OnOff.ToString.ToLower Then
                    Return "On"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.OneZero.ToString.ToLower Then
                    Return "One"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.TrueFalse.ToString.ToLower Then
                    Return "True"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.IfNot.ToString.ToLower Then
                    Return "If possible"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.SucceededFailed.ToString.ToLower Then
                    Return "Succeeded"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.Cart.ToString.ToLower Then
                    Return "Yes, Include This Item"
                End If
            Case False
                If pattern_.ToString.ToLower = DropTextPattern.AlwaysNever.ToString.ToLower Then
                    Return "Never"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.YesNo.ToString.ToLower Then
                    Return "No"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.OnOff.ToString.ToLower Then
                    Return "Off"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.OneZero.ToString.ToLower Then
                    Return "Zero"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.TrueFalse.ToString.ToLower Then
                    Return "False"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.IfNot.ToString.ToLower Then
                    Return "Not at all"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.SucceededFailed.ToString.ToLower Then
                    Return "Failed"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.Cart.ToString.ToLower Then
                    Return "No, Remove This Item"
                End If
        End Select
    End Function
    ''' <summary>
    ''' Converts user-friendly word like yes or always to boolean i.e. true or false.
    ''' To convert back to user-friendly word, use BooleanToDropText(boolean_val As Boolean, Optional pattern_ As String = "always/never")
    ''' </summary>
    ''' <example>
    ''' <code>
    ''' Dim userDecision as string = DropTextToBoolean(dropText.text)
    ''' </code>
    ''' </example>
    ''' <param name="str_">always/never (default), yes/no, on/off, 1/0, true/false, if/not</param>
    ''' <returns>True or False</returns>
    Public Shared Function DropTextToBoolean(str_ As String) As Boolean
        If BooleanDropTextTrue.Contains(str_.ToLower) Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region

    Public Function Toboolean(str_)
        Convert.ToBoolean(Convert.ToInt32(str_))
    End Function


End Class

Public Class ManagedPower
    ' GetSystemPowerStatus() is the only unmanaged API called.
    Declare Auto Function GetSystemPowerStatus Lib "kernel32.dll" _
    Alias "GetSystemPowerStatus" (ByRef sps As SystemPowerStatus) As Boolean

    Public Overrides Function ToString() As String
        Dim text As String = ""
        Dim sysPowerStatus As SystemPowerStatus
        ' Get the power status of the system
        If ManagedPower.GetSystemPowerStatus(sysPowerStatus) Then
            ' Current power source - AC/DC
            Dim currentPowerStatus = sysPowerStatus.ACLineStatus
            text += sysPowerStatus.ACLineStatus.ToString()
        End If
        Return text
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure SystemPowerStatus
        Public ACLineStatus As _ACLineStatus
        Public BatteryFlag As _BatteryFlag
        Public BatteryLifePercent As Byte
        Public Reserved1 As Byte
        Public BatteryLifeTime As System.UInt32
        Public BatteryFullLifeTime As System.UInt32
    End Structure

    Public Enum _ACLineStatus As Byte
        Battery = 0
        AC = 1
        Unknown = 255
    End Enum

    <Flags()>
    Public Enum _BatteryFlag As Byte
        High = 1
        Low = 2
        Critical = 4
        Charging = 8
        NoSystemBattery = 128
        Unknown = 255
    End Enum
End Class


