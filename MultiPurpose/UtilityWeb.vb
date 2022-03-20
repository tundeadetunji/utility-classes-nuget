Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls
Imports MultiPurpose.Utility
Public Class UtilityWeb
#Region ""
    Public Enum ShowElementAfterward
        Yes = 1
        No = 2
    End Enum
    Public Enum AlertIs
        alert
        danger
        success
        warning
    End Enum
    Public Enum Fix
        LineBreak
        LineBreakWithParagraph
    End Enum
    Enum WriteContentType
        None
        Fix
        Bootstrap
        Jumbotron
    End Enum
    Enum ControlsToCheck
        All
        Any
    End Enum

#End Region

    Public Shared Sub EnableControl(div_ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional state_ As Boolean = True)
        div_.Visible = state_
    End Sub

    Public Shared Sub Clear(c_ As Array)
        For i As Integer = 0 To c_.Length - 1
            Clear(c_(i), "")
        Next
    End Sub
    Public Shared Sub Clear(c_ As Object, Optional default_string As String = "")

        Dim c As DropDownList
        Dim l As ListBox
        Dim t As TextBox
        Dim p As System.Web.UI.WebControls.Image
        Dim h As CheckBox
        Dim ht As HtmlInputText
        Dim hta As HtmlTextArea
        Dim html_select As HtmlSelect
        Dim g As GridView

        If TypeOf c_ Is GridView Then
            g = c_
            Try
                g.DataSource = Nothing
            Catch
            End Try
        End If

        If TypeOf c_ Is CheckBox Then
            h = c_
            h.Checked = False
        End If
        If TypeOf c_ Is DropDownList Then
            c = c_
            Try
                c.DataSource = Nothing
            Catch ex As Exception
            End Try
            c.Items.Clear()
            c.Text = ""
        End If
        If TypeOf c_ Is ListBox Then
            l = c_
            Try
                l.DataSource = Nothing
            Catch ex As Exception
            End Try
            l.Items.Clear()
        End If
        If TypeOf c_ Is TextBox Then
            t = c_
            t.Text = default_string
        End If
        If TypeOf c_ Is System.Web.UI.WebControls.Image Then
            p = c_
            Try
                p.ImageUrl = Nothing
            Catch ex As Exception
            End Try
        End If
        If TypeOf c_ Is HtmlInputText Then
            ht = c_
            ht.Value = default_string
        End If
        If TypeOf c_ Is HtmlTextArea Then
            hta = c_
            hta.Value = default_string
        End If
        If TypeOf c_ Is HtmlSelect Then
            html_select = c_
            Try
                html_select.Items.Clear()
                html_select.Value = default_string
            Catch ex As Exception
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Places text inside TextBox, ComboBox, Button, HTML DIV/SPAN or Label, converts date to short (and also for web if DIV/SPAN) alongside.
    ''' </summary>
    ''' <param name="str_"></param>
    ''' <param name="control_"></param>
    ''' <param name="convert_date_to_short"></param>
    ''' <returns></returns>
    Public Shared Function Write(str_ As Object, control_ As Object, Optional convert_date_to_short As Boolean = True) As String
        Dim r As String '= DateToShort(str_, convert_date_to_short)
        If convert_date_to_short Then
            r = Date.Parse(str_).ToShortDateString
        Else
            r = str_.ToString
        End If
        Dim t As TextBox, d As DropDownList, l As Label, b As Button, div_ As System.Web.UI.HtmlControls.HtmlGenericControl
        Try
            If TypeOf control_ Is TextBox Then
                t = control_
                t.Text = r
            ElseIf TypeOf control_ Is DropDownList Then
                d = control_
                d.Text = r
            ElseIf TypeOf control_ Is Label Then
                l = control_
                l.Text = r
            ElseIf TypeOf control_ Is Button Then
                b = control_
                b.Text = r
            ElseIf TypeOf control_ Is System.Web.UI.HtmlControls.HtmlGenericControl Then
                div_ = control_
                WriteContent(PrepareForIO(r), div_)
            End If
        Catch ex As Exception

        End Try
        Return r
    End Function
    Public Shared Function Jumbotron(str__ As String, div__ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional size__ As Byte = 3) As String
        Dim size As Byte = 3
        If size__ < 1 Or size__ > 7 Then size = 7
        Dim str_ As String =
            "<div class=""jumbotron"">
                <h" & size & ">" & str__ & "</h" & size & ">
            </div>"
        div__.InnerHtml = str_
        Return str_
    End Function
    ''' <summary>
    ''' Creates empty space.
    ''' </summary>
    ''' <param name="length_">lt 0 for 2 Line breaks.</param>
    ''' <returns></returns>
    Public Shared Function fix_str(Optional length_ As Integer = 5, Optional fix_ As Fix = Fix.LineBreak) As String
        If length_ < 0 Then Return "<br /><br />"

        Dim line_break As String = "<br /><br />"
        Dim paragraph_ As String = "<p>&nbsp;<br />&nbsp;</p>"

        Dim r_ As String
        Select Case fix_
            Case Fix.LineBreak
                r_ = line_break
            Case Fix.LineBreakWithParagraph
                r_ = paragraph_
        End Select

        Dim r As String
        For i As Integer = 1 To length_
            r &= r_
        Next
        Return r
    End Function

    Public Shared Sub WriteContent(div_ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional length_ As Integer = 1, Optional content_type As Fix = Fix.LineBreak)
        WriteContent(fix_str(length_), div_, content_type)
    End Sub

    Public Shared Function WriteContent(div_ As System.Web.UI.HtmlControls.HtmlGenericControl, str_ As String, Optional write_content As WriteContentType = WriteContentType.None, Optional WithClose As Boolean = True, Optional alert_is As AlertIs = AlertIs.warning, Optional format_for_web As Boolean = True)
        Return WriteContent(str_, div_, write_content, WithClose, alert_is, format_for_web)
    End Function
    Public Shared Function WriteContent(str_ As String, div_ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional write_content As WriteContentType = WriteContentType.None, Optional WithClose As Boolean = True, Optional alert_is As AlertIs = AlertIs.warning, Optional format_for_web As Boolean = True)
        If write_content = WriteContentType.Jumbotron Then Return Jumbotron(str_, div_)
        Dim fix_ As String = fix_str()
        Dim str__ As String = str_
        If format_for_web Then str_ = ToIO(str_)

        Try
            If write_content = WriteContentType.Bootstrap Then
                Alert(str__, div_, WithClose, alert_is)
            ElseIf write_content = WriteContentType.None Then
                div_.InnerHtml = str__
                div_.Visible = True
            ElseIf write_content = WriteContentType.Fix Then
                div_.InnerHtml = fix_
                div_.Visible = True
            End If
        Catch ex As Exception
        End Try
    End Function
    ''' <summary>
    ''' Gives feedback to user. Same as Feedback. Bootstrap 3.
    ''' </summary>
    ''' <param name="str_"></param>
    ''' <param name="WithClose"></param>
    ''' <param name="alert_is"></param>
    ''' <returns></returns>
    Public Shared Function Alert(str_ As String, div As System.Web.UI.HtmlControls.HtmlGenericControl, Optional WithClose As Boolean = True, Optional alert_is As AlertIs = AlertIs.warning) As String
        Try
            Dim r As String = Feedback(str_, WithClose, alert_is.ToString)
            div.InnerHtml = r
            If ShowElementAfterAlert = ShowElementAfterward.Yes Then
                div.Visible = True
            Else
                div.Visible = False
            End If
            Return r
        Catch
        End Try
    End Function
    Public Shared Property ShowElementAfterAlert As ShowElementAfterward = ShowElementAfterward.Yes
    Public Shared Function FeedbackWithClose(str_ As String, alert_OR_danger_OR_success_OR_warning As String) As String
        Dim header_ As String = "", footer_ As String = "</div>"
        Dim close_ As String = "<button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close""><span aria-hidden=""true"">&times;</span></button>"

        Select Case alert_OR_danger_OR_success_OR_warning.ToLower
            Case "alert"
                header_ = "<div class=""alert alert-primary alert-dismissible fade show"" role=""alert"">"
            Case "danger"
                header_ = "<div class=""alert alert-danger alert-dismissible fade show"" role=""alert"">"
            Case "success"
                header_ = "<div class=""alert alert-success alert-dismissible fade show"" role=""alert"">"
            Case "warning"
                header_ = "<div class=""alert alert-warning alert-dismissible fade show"" role=""alert"">"
        End Select

        Return header_ & str_ & close_ & footer_

    End Function
    ''' <summary>
    ''' Gives feedback to user. Constructed as an alert DIV.
    ''' </summary>
    ''' <param name="str_"></param>
    ''' <param name="WithClose"></param>
    ''' <param name="alert_OR_danger_OR_success_OR_warning"></param>
    ''' <returns></returns>
    Public Shared Function Feedback(str_ As String, Optional WithClose As Boolean = True, Optional alert_OR_danger_OR_success_OR_warning As String = "warning") As String
        Select Case WithClose
            Case True
                Return FeedbackWithClose(str_, alert_OR_danger_OR_success_OR_warning)
            Case False
                Return FeedbackWithoutClose(str_, alert_OR_danger_OR_success_OR_warning)
        End Select



        '<span runat = "server" id="x"></span>
        'x.InnerHtml = f.Feedback(f.InvalidCredentialFeedback, True, "alert")
        'x.Visible = True

    End Function
    Public Shared Function FeedbackWithoutClose(str_ As String, alert_OR_danger_OR_success_OR_warning As String) As String
        Dim header_ As String = "", footer_ As String = "</div>"
        Select Case alert_OR_danger_OR_success_OR_warning.ToLower
            Case "alert"
                header_ = "<div class=""alert alert-primary text-center"" role=""alert"">"
            Case "danger"
                header_ = "<div class=""alert alert-danger text-center"" role=""alert"">"
            Case "success"
                header_ = "<div class=""alert alert-success text-center"" role=""alert"">"
            Case "warning"
                header_ = "<div class=""alert alert-warning text-center"" role=""alert"">"
        End Select
        Return header_ & str_ & footer_
    End Function
    Public Shared Function ToIO(str_ As String, Optional output_ As FormatFor = FormatFor.Web) As String
        Return PrepareForIO(str_, output_)
    End Function

    Public Shared Function Content(f As FileUpload) As Object
        If TypeOf f Is FileUpload Then Return CType(f, FileUpload).FileBytes
    End Function

    Public Shared Property casing__ As TextCase = TextCase.Capitalize
    Public Shared Function Content(control_ As Object, Optional casing_ As TextCase = TextCase.None) As String
        Dim casing As TextCase = casing__
        If casing_ <> Nothing Then casing = casing_

        Try
            If TypeOf control_ Is TextBox Then
                Return TransformText(CType(control_, TextBox).Text, casing)
            ElseIf TypeOf control_ Is HtmlInputText Then
                Return TransformText(CType(control_, HtmlInputText).Value, casing)
            ElseIf TypeOf control_ Is HtmlTextArea Then
                Return CType(control_, HtmlTextArea).Value
            ElseIf TypeOf control_ Is DropDownList Then
                Return TransformText(CType(control_, DropDownList).Text, casing)
            ElseIf TypeOf control_ Is HtmlSelect Then
                Return TransformText(CType(control_, HtmlSelect).Items.Item(control_.SelectedIndex).ToString, casing)
            ElseIf TypeOf control_ Is CheckBox Then
                Return CType(control_, CheckBox).Checked
            End If
        Catch
        End Try
    End Function

    Public Shared Sub TitleDrop(c_ As DropDownList, Optional first_item_is_empty As Boolean = True)
        If c_.Items.Count > 0 Then Exit Sub

        Clear(c_)
        With c_
            With .Items
                If first_item_is_empty Then .Add("")
                .Add("Mr.")
                .Add("Mrs.")
                .Add("Ms.")
            End With
            .SelectedIndex = -1
            .Text = ""
        End With

    End Sub

    ''' <summary>
    ''' Determines if all controls have text.
    ''' </summary>
    ''' <param name="controls_">Controls</param>
    ''' <returns>True or False</returns>

    Private Shared Function IsEmpty(controls_ As Array, Optional controls_to_check As ControlsToCheck = ControlsToCheck.Any) As Boolean
        Dim counter_ As Integer = 0
        With controls_
            For i As Integer = 0 To .Length - 1
                If IsEmpty(controls_(i)) Then
                    counter_ += 1
                End If
            Next
        End With
        '		Return Val(counter_) = controls_.Length
        If controls_to_check = ControlsToCheck.All Then
            Return Val(counter_) = controls_.Length
        Else
            Return Val(counter_) > 0
        End If
    End Function

    ''' <summary>
    ''' Determines if control has text or list has items.
    ''' </summary>
    ''' <param name="c_"></param>
    ''' <param name="use_trim">Should content be trimmed before check?</param>
    ''' <returns></returns>
    Public Shared Function IsEmpty(c_ As Object, Optional use_trim As Boolean = True) As Boolean
        Dim d__ As DropDownList
        Dim t__ As TextBox
        Dim html_inputText As HtmlInputText
        Dim html_textAra As HtmlTextArea
        Dim html_select As HtmlSelect
        Dim a As Array
        Dim l_string As List(Of String)
        Dim l_object As List(Of Object)
        Dim l_integer As List(Of Integer)
        Dim l_double As List(Of Double)
        Dim c As Collection

        If TypeOf c_ Is Array Then
            a = c_
            Return a.Length < 1
        End If

        If TypeOf c_ Is List(Of String) Then
            l_string = c_
            Return l_string.Count < 1
        End If

        If TypeOf c_ Is List(Of Object) Then
            l_object = c_
            Return l_object.Count < 1
        End If

        If TypeOf c_ Is List(Of Integer) Then
            l_integer = c_
            Return l_integer.Count < 1
        End If

        If TypeOf c_ Is List(Of Double) Then
            l_double = c_
            Return l_double.Count < 1
        End If

        If TypeOf c_ Is Collection Then
            c = c_
            Return c.Count < 1
        End If

        If TypeOf c_ Is DropDownList Then
            d__ = c_
            If use_trim = True Then
                Return d__.Text.Trim.Length < 1
            Else
                Return d__.Text.Length < 1
            End If
        End If

        If TypeOf c_ Is TextBox Then
            t__ = c_
            If use_trim = True Then
                Return t__.Text.Trim.Length < 1
            Else
                Return t__.Text.Length < 1
            End If
        End If

        If TypeOf c_ Is HtmlInputText Then
            html_inputText = c_
            If use_trim = True Then
                Return html_inputText.Value.Trim.Length < 1
            Else
                Return html_inputText.Value.Length < 1
            End If
        End If

        If TypeOf c_ Is HtmlTextArea Then
            html_textAra = c_
            If use_trim = True Then
                Return html_textAra.Value.Trim.Length < 1
            Else
                Return html_textAra.Value.Length < 1
            End If
        End If

        If TypeOf c_ Is HtmlSelect Then
            html_select = c_
            Return html_select.Items.Item(html_select.SelectedIndex).ToString.Length < 1
        End If

        If TypeOf (c_) Is FileUpload Then
            If CType(c_, FileUpload).HasFile = False Or CType(c_, FileUpload).HasFiles = False Then
                Return True
            Else
                Return False
            End If
        End If

        If TypeOf c_ Is CheckBox Then Return CType(c_, CheckBox).Checked = False

    End Function
    ''' <summary>
    ''' Populates drop with drop-text version of boolean, depending on the pattern.
    ''' </summary>
    ''' <param name="d_">DropDownList</param>
    ''' <param name="firstItemIsEmpty">Should DropDownList's first item be empty?</param>
    ''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
    Private Shared Sub BooleanDrop(d_ As DropDownList, Optional pattern_ As String = "always/never", Optional firstItemIsEmpty As Boolean = False)
        '		Dim format_window_web As New NFunctions
        With d_
            If .Items.Count > 0 Then Exit Sub
            With .Items
                If firstItemIsEmpty = True Then .Add("")
                .Add(BooleanToDropText(True, pattern_))
                .Add(BooleanToDropText(False, pattern_))
            End With
        End With
    End Sub

    Public Shared Sub BooleanDrop(d_ As DropDownList, Optional pattern_ As DropTextPattern = DropTextPattern.AlwaysNever, Optional firstItemIsEmpty As Boolean = False)
        With d_
            If .Items.Count > 0 Then Exit Sub
            With .Items
                If firstItemIsEmpty = True Then .Add("")
                .Add(BooleanToDropText(True, pattern_))
                .Add(BooleanToDropText(False, pattern_))
            End With
            Try
                .Text = ""
                .SelectedIndex = -1
            Catch ex As Exception

            End Try
        End With

    End Sub





End Class
