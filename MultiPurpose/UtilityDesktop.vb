Imports System.Drawing
Imports System.Windows.Forms
Imports MultiPurpose.Utility
Public Class UtilityDesktop

    Private suffx
    Private prefx
    Private countr
    Private TextHasSpace As Boolean
    Private strSource

#Region ""
    Public Enum TimeValue
        ShortDate = 1
        LongDate = 2
        ShortTime = 3
        LongTime = 4
        Year = 5
        Month = 6
        Day = 7
        Hour = 8
        Minute = 9
        Second = 10
        Millisecond = 11
        DayOfWeek = 12
        DayOfYear = 13
    End Enum
    Public Enum ReturnInfo
        AsArray = 0
        AsString = 1
        AsListOfString = 2
        AsListOfObject = 3
        AsCollectionOfString = 4
        AsCollectionOfObject = 5
    End Enum

    Enum ControlsToCheck
        All
        Any
    End Enum

#End Region

    ''' <summary>
    ''' Changes text to title case. Called from _TextChanged. Works for single-line textboxes.
    ''' </summary>
    ''' <param name="strSource"></param>
    Public Sub ConvertTextToTitleCase(ByRef strSource As Control)

        'convert text to title case
        'called from TextChanged
        If TypeOf strSource Is TextBox Then
            Dim t As TextBox = strSource
            If t.Multiline = True Then Exit Sub
        End If

        On Error Resume Next
        If Len(strSource.Text) = 0 Then
            suffx = ""
            prefx = ""
            countr = Nothing
            Exit Sub
        End If

        If Len(strSource.Text) = 1 Then strSource.Text = UCase(strSource.Text) : System.Windows.Forms.SendKeys.Send("{End}")

        If Mid(strSource.Text, Len(strSource.Text), 1) = " " Or Mid(strSource.Text, Len(strSource.Text), 1) = "." Or Mid(strSource.Text, Len(strSource.Text), 1) = ChrW(13) Then
            TextHasSpace = True
            countr = Len(strSource.Text) + 1
            prefx = strSource.Text
        End If

        If prefx <> "" And Len(strSource.Text) = Val(countr) Then
            suffx = UCase(Mid(strSource.Text, Len(strSource.Text), 1))
            strSource.Text = prefx & suffx
            System.Windows.Forms.SendKeys.Send("{End}")
            'clear counters
            suffx = ""
            prefx = ""
            countr = Nothing
        End If
    End Sub

    ''' <summary>
    ''' Checks if Enter key was pressed. Called from _KeyPress.
    ''' </summary>
    ''' <param name="e"></param>
    ''' <returns></returns>

    Public Shared Function EnterKeyWasPressed(ByRef e As System.Windows.Forms.KeyPressEventArgs) As Boolean
        Return e.KeyChar = ChrW(13)
    End Function

    ''' <summary>
    ''' Allows numbers and % only. Called from _KeyPress.
    ''' </summary>
    ''' <param name="e"></param>
    ''' <param name="controlText"></param>
    Public Shared Sub AllowPercentage(ByRef e As System.Windows.Forms.KeyPressEventArgs, controlText As String)
        '        On Error Resume Next
        'allow numbers(48-57), backspace(8), enter(13), full stop(46), delete(127), percent(37) and minus(45) only
        If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Or e.KeyChar = ChrW(46) Or e.KeyChar = ChrW(45) Or e.KeyChar = ChrW(37) Then
            Exit Sub
        Else
            e.KeyChar = ChrW(0)
        End If
    End Sub

    ''' <summary>
    ''' Allows positive and negative numbers and period only. Called from _KeyPress. Good for money.
    ''' </summary>
    ''' <param name="e"></param>
    Public Shared Sub AllowNumberOnly(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        '        On Error Resume Next
        'allow numbers(48-57), backspace(8), enter(13), full stop(46), delete(127) and minus(45) only
        If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Or e.KeyChar = ChrW(46) Or e.KeyChar = ChrW(45) Then
            Exit Sub
        Else
            e.KeyChar = ChrW(0)
        End If
    End Sub
    Public Shared Sub AllowFileNameCharacterOnly(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        'alphabets, numbers, underscore, backspace
        If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or (e.KeyChar >= ChrW(65) And e.KeyChar <= ChrW(90)) Or (e.KeyChar >= ChrW(97) And e.KeyChar <= ChrW(122)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(95) Then
        Else
            e.KeyChar = ChrW(0)
        End If
    End Sub

    ''' <summary>
    ''' Allows numbers only. Called from _KeyPress. Good for PIN.
    ''' </summary>
    ''' <param name="e"></param>

    Public Shared Sub AllowDigitOnly(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        '        On Error Resume Next
        'allow numbers(48-57), backspace(8), enter(13), full stop(46), delete(127) only
        If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Or e.KeyChar = ChrW(46) Then
            Exit Sub
        Else
            e.KeyChar = ChrW(0)
        End If
    End Sub

    ''' <summary>
    ''' Allows nothing.
    ''' </summary>
    ''' <param name="e"></param>
    Public Shared Sub AllowNothing(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        e.KeyChar = ChrW(0)
    End Sub
    ''' <summary>
    ''' Changes text to uppercase. Called from _TextChanged.
    ''' </summary>
    ''' <param name="strSource"></param>
    Public Shared Sub ToUpperCase(ByRef strSource As Control)
        '		On Error Resume Next
        'called from TextChanged
        If Len(strSource.Text) = 0 Then Exit Sub
        Try
            strSource.Text = UCase(strSource.Text)
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Changes text to lower case. Called from _TextChanged. Useful for email addresses and websites.
    ''' </summary>
    ''' <param name="strSource"></param>
    Public Sub ToLowerCase(ByRef strSource As Control)
        On Error Resume Next
        'convert text to lower case (useful for email and websites addresses)
        'called from TextChanged
        If Len(strSource.Text) = 0 Then Exit Sub

        strSource.Text = LCase(strSource.Text)

    End Sub

    ''' <summary>
    ''' Search.
    ''' </summary>
    ''' <param name="control_">Textbox to search from.</param>
    ''' <param name="sought_">What to select if found in control_'s content.</param>
    ''' <param name="current_position">Variable to keep track of how many of sought_ has been found (should be overwritten at each next call otherwise it'll only always check once).</param>
    ''' <returns>0 if sought_ isn't found, else next index where it is found.</returns>
    ''' <example>
    ''' 'DECLARATION
    ''' Public search_position As Integer = 1
    ''' Public search_ As String = ""
    ''' 'Ctrl+F
    ''' search_ = GetText("Locate ...")
    ''' search_position = LocateText(tContent, search_, search_position)
    ''' 'F3
    ''' If search_.Length > 0 Then search_position = LocateText(tContent, search_, search_position)
    ''' </example>
    Public Shared Function LocateText(control_ As TextBox, sought_ As String, Optional current_position As Integer = Nothing) As Integer
        If control_.Text.Length < 1 Then Return 0
        Dim search_ As Integer = 1
        If current_position <> Nothing Then search_ = Val(current_position)
        If search_ = 0 Then search_ = 1

        For i% = search_ To control_.Text.Length
            If Mid(control_.Text, i, sought_.Length).ToLower() = sought_.Trim.ToLower() Then
                control_.SelectionStart = i - 1
                control_.SelectionLength = sought_.Length
                control_.ScrollToCaret()
                control_.Focus()
                search_ = i + 1
                Exit For
            ElseIf Mid(control_.Text, i, sought_.Length).ToLower() <> sought_.Trim.ToLower() And control_.Text.Length - i < sought_.Length Then
                search_ = 0
                Exit For
            End If
        Next
        Return search_

        'DECLARATION
        'Public search_position As Integer = 1
        'Public search_ As String = ""
        'Ctrl+F
        'search_ = GetText("Locate ...")
        'search_position = LocateText(tContent, search_, search_position)
        'F3
        'If search_.Length > 0 Then search_position = LocateText(tContent, search_, search_position)


        'INITIALLY
        'u.text = 0
        'RECURRENT, IF u.text > 0
        'u.Text = LocateText(t, "html", u.Text)

    End Function

    Public Shared Sub Sleep()
        Try
            Application.SetSuspendState(PowerState.Suspend, True, False)
        Catch ex As Exception

        End Try
    End Sub
    Public Shared Sub Hibernate()
        Try
            Application.SetSuspendState(PowerState.Hibernate, True, True)
        Catch
        End Try
    End Sub

#Region "Sound"
    Public Enum MakeTheFile
        Hidden
        System
        Invisible
        Visible
    End Enum

    Private Declare Function record Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
    Private Property filename__recording As String
    Public Property elapsed_label__recording As Label
    Public Property remaining_label__recording As Label
    Public Property append_text__recording As Boolean
    Public Property what_to_display_when_finished__recording As String
    Private Property duration_in_seconds__recording As Long
    Private Property make_the_file__recording As MakeTheFile
    Private timer_recording__ As Timer
    Private counter__recording As Long
    Public Structure DisplayControlsForRecordingDuration
        Public elapsed_label As Label
        Public remaining_label As Label
        Public append_text As Boolean
        Public what_to_display_when_finished As String
    End Structure
    ''' <summary>
    ''' Starts an audio recording.
    ''' </summary>
    ''' <param name="directory_"></param>
    ''' <param name="file_name_without_extension"></param>
    ''' <param name="timer_"></param>
    ''' <param name="duration_in_seconds"></param>
    ''' <param name="file_extension"></param>
    ''' <param name="display_labels"></param>
    ''' <param name="auto_name_file"></param>
    Public Sub Recording(directory_ As String, file_name_without_extension As String, Optional timer_ As Timer = Nothing, Optional duration_in_seconds As Long = 60, Optional file_extension As String = ".wav", Optional display_labels As DisplayControlsForRecordingDuration = Nothing, Optional auto_name_file As Boolean = False)
        If directory_.Length < 1 Then Exit Sub
        If file_name_without_extension.Length < 1 And auto_name_file = False Then Exit Sub

        If timer_ IsNot Nothing Then timer_.Enabled = False

        Try
            My.Computer.FileSystem.CreateDirectory(directory_)
        Catch ex As Exception
        End Try

        Dim filename_ As String = directory_ & "\"
        If auto_name_file = True Then
            filename_ &= String.Format("{0:00}", CStr(Now.Month)) & "." & String.Format("{0:00}", CStr(Now.Day)) & "." & String.Format("{0:0000}", CStr(Now.Year)) & "T" & String.Format("{0:00}", CStr(Now.Hour)) & "." & String.Format("{0:00}", CStr(Now.Minute)) & "." & String.Format("{0:00}", CStr(Now.Millisecond)) & file_extension
        Else
            filename_ &= file_name_without_extension & file_extension
        End If
        filename__recording = filename_

        If Not IsNothing(display_labels) Then
            elapsed_label__recording = display_labels.elapsed_label
            remaining_label__recording = display_labels.remaining_label
            append_text__recording = display_labels.append_text
            what_to_display_when_finished__recording = display_labels.what_to_display_when_finished
        End If

        Dim duration_ As Long = 60000
        If duration_in_seconds <> Nothing Then duration_ = duration_in_seconds
        If timer_ IsNot Nothing Then
            timer_recording__ = timer_
            AddHandler timer_.Tick, New EventHandler(AddressOf RecordingTimer)
            timer_.Interval = Val(duration_) * 1000 '* 60
            timer_.Enabled = True
        End If

        counter__recording = 0
        record("open new Type waveaudio Alias recsound", "", 0, 0)
        record("record recsound", "", 0, 0)

    End Sub
    ''' <summary>
    ''' Starts an audio recording.
    ''' </summary>
    ''' <param name="directory_"></param>
    ''' <param name="file_name_without_extension"></param>
    ''' <param name="timer_"></param>
    ''' <param name="duration_in_seconds"></param>
    ''' <param name="file_extension"></param>
    ''' <param name="display_labels"></param>
    ''' <param name="auto_name_file"></param>
    Public Sub Recording(directory_ As String, file_name_without_extension As String, Optional timer_ As Timer = Nothing, Optional duration_in_seconds As Long = 60, Optional file_extension As String = ".wav", Optional display_labels As DisplayControlsForRecordingDuration = Nothing, Optional auto_name_file As Boolean = False, Optional make_the_file As MakeTheFile = MakeTheFile.Visible)
        If directory_.Length < 1 Then Exit Sub
        If file_name_without_extension.Length < 1 And auto_name_file = False Then Exit Sub

        make_the_file__recording = make_the_file

        Recording(directory_, file_name_without_extension, timer_, duration_in_seconds, file_extension, display_labels, auto_name_file)
    End Sub
    ''' <summary>
    ''' Ends the recording started with Recording()
    ''' </summary>
    Public Sub EndRecording()
        record("save recsound " & filename__recording, "", 0, 0)
        record("close recsound", "", 0, 0)
        Select Case make_the_file__recording
            Case MakeTheFile.Hidden
                SetAttr(filename__recording, FileAttribute.Hidden)
            Case MakeTheFile.Invisible
                SetAttr(filename__recording, FileAttribute.Hidden + FileAttribute.System)
            Case MakeTheFile.System
                SetAttr(filename__recording, FileAttribute.System)
            Case MakeTheFile.Visible
                SetAttr(filename__recording, FileAttribute.Normal)
        End Select
    End Sub

    Private Sub RecordingTimer()
        If counter__recording = duration_in_seconds__recording Then
            timer_recording__.Enabled = False
            EndRecording()
            Exit Sub
        End If
        Dim remaining As Long = duration_in_seconds__recording - counter__recording
        If elapsed_label__recording IsNot Nothing Then
            If append_text__recording Then
                elapsed_label__recording.Text = counter__recording & " seconds into recording"
            Else
                elapsed_label__recording.Text = counter__recording
            End If
        End If
        If remaining_label__recording IsNot Nothing Then
            If append_text__recording Then
                remaining_label__recording.Text = remaining & " seconds remaining"
            Else
                remaining_label__recording.Text = remaining
            End If
        End If
        counter__recording += 1
    End Sub

#End Region

    ''' <summary>
    ''' Grabs all items in listbox and returns them in as List(of String) or Array
    ''' </summary>
    ''' <param name="l"></param>
    ''' <param name="return_as"></param>
    ''' <returns></returns>
    Public Shared Function ListToArray(l As ListBox, return_as As ReturnInfo) As Object
        If l Is Nothing Then Return Nothing
        If l.Items.Count < 1 Then Return Nothing
        Dim l_ As New List(Of String)
        With l.Items
            For i As Integer = 0 To .Count - 1
                l_.Add(.Item(i).ToString)
            Next
        End With
        Select Case return_as
            Case ReturnInfo.AsListOfString
                Return l_
            Case Else
                Return l_.ToArray
        End Select
    End Function
    ''' <summary>
    ''' Grabs all items in combobox and returns them in as List(of String) or Array
    ''' </summary>
    ''' <param name="l"></param>
    ''' <param name="return_as"></param>
    ''' <returns></returns>
    Public Shared Function ListToArray(l As ComboBox, return_as As ReturnInfo) As Object
        If l Is Nothing Then Return Nothing
        If l.Items.Count < 1 Then Return Nothing
        Dim l_ As New List(Of String)
        With l.Items
            For i As Integer = 0 To .Count - 1
                l_.Add(.Item(i).ToString)
            Next
        End With
        Select Case return_as
            Case ReturnInfo.AsListOfString
                Return l_
            Case Else
                Return l_.ToArray
        End Select
    End Function

    Public Shared Sub EnableControls(control_ As Array, Optional state_ As Boolean = True)
        On Error Resume Next
        If control_.Length < 1 Then Exit Sub
        For i As Integer = 0 To control_.Length - 1
            control_(i).Enabled = state_
        Next

    End Sub

    Public Shared Sub EnableControl(control As Control, Optional state_ As Boolean = True)
        On Error Resume Next
        control.Enabled = state_

    End Sub

    Public Shared Sub DisableControl(control As Control)
        EnableControl(control, False)

    End Sub

    ''' <summary>
    ''' Determines if all controls and/or file(s) do not have text.
    ''' </summary>
    ''' <param name="controls_">Controls or File paths or both</param>
    ''' <returns>True or False</returns>

    Public Shared Function IsEmpty(controls_ As Array, controls_to_check As ControlsToCheck) As Boolean
        Dim counter_ As Integer = 0
        Dim control_to_focus As Control
        With controls_
            For i As Integer = 0 To .Length - 1
                If IsEmpty(controls_(i), True, Nothing) Then
                    control_to_focus = controls_(i)
                    counter_ += 1
                End If
            Next
        End With
        If controls_to_check = ControlsToCheck.All Then
            Return Val(counter_) = controls_.Length
        Else
            Return Val(counter_) > 0
        End If
    End Function

    ''' <summary>
    ''' Determines if control or file does not have text (or items, if it's listbox, array, list of string or object or integer or double, collection).
    ''' </summary>
    ''' <param name="c_">File path or ComboBox or TextBox or PictureBox or NumericUpDown</param>
    ''' <param name="use_trim">Should content be trimmed before check?</param>
    ''' <param name="control_to_focus">Focus on this when IsEmpty is True</param>
    ''' <returns>True or False</returns>
    Public Shared Function IsEmpty(c_ As Object, Optional use_trim As Boolean = True, Optional control_to_focus As Control = Nothing) As Boolean
        Dim n__ As NumericUpDown
        Dim p__ As PictureBox
        ''        Dim d__ As ComboBox
        Dim t__ As TextBox
        Dim l__ As ListBox

        Dim return_val As Boolean = False

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

        If TypeOf c_ Is PictureBox Then
            p__ = c_
            If p__.BackgroundImage Is Nothing And p__.Image Is Nothing Then
                return_val = True
            End If
        ElseIf TypeOf c_ Is NumericUpDown Then
            n__ = c_
            If (n__.Value) = 0 Then
                return_val = True
            End If
        ElseIf TypeOf c_ Is ListBox Then
            l__ = c_
            If l__.Items.Count < 1 Then
                return_val = True
            End If
        ElseIf TypeOf c_ Is ComboBox Then
            If use_trim = True Then
                If CType(c_, ComboBox).Text.Trim.Length < 1 Then
                    return_val = True
                End If
            Else
                If CType(c_, ComboBox).Text.Length < 1 Then
                    return_val = True
                End If
            End If
        ElseIf TypeOf c_ Is TextBox Then
            t__ = c_
            If use_trim = True Then
                If t__.Text.Trim.Length < 1 Then
                    return_val = True
                End If
            Else
                If t__.Text.Length < 1 Then
                    return_val = True
                End If
            End If
        ElseIf TypeOf c_ Is CheckBox Then
            Return CType(c_, CheckBox).Checked = False
        ElseIf CType(c_, ComboBox).Text.Trim.Length < 1 Then
            return_val = True
        Else
            Try
                If ReadText(c_).Length < 1 Then
                    return_val = True
                End If
            Catch ex As Exception
            End Try
        End If
        'If return_val = True And string_feedback.Length > 0 And app_.Length > 0 Then
        '    Try
        '        Dim feedback_ As New Feedback.Feedback()
        '        feedback_.fFeedback(string_feedback)
        '    Catch
        '    End Try
        'End If
        If control_to_focus IsNot Nothing Then
            Try
                control_to_focus.Focus()
            Catch ex As Exception
            End Try
        ElseIf control_to_focus Is Nothing Then
            Try
                c_.Focus()
            Catch ex As Exception
            End Try
        End If
        Return return_val
    End Function

    Public Shared Sub TitleDrop(c_ As ComboBox, Optional first_item_is_empty As Boolean = False)
        If c_.Items.Count > 0 Then Exit Sub

        Clear(c_)
        With c_
            With .Items
                If first_item_is_empty Then .Add("")
                .Add("Mr.")
                .Add("Mrs.")
                .Add("Ms.")
            End With
            .Sorted = True
            .SelectedIndex = -1
            .Text = ""
        End With

    End Sub
    Public Shared Function PictureFromStream(picture_ As Object, Optional file_extension As String = ".jpg", Optional UseImage As Boolean = False)
        Dim photo_ 'As Image
        Dim stream_ As New IO.MemoryStream

        If TypeOf picture_ Is PictureBox Then
            Select Case UseImage
                Case True
                    photo_ = picture_.Image
                Case False
                    photo_ = picture_.BackgroundImage
            End Select
        ElseIf TypeOf picture_ Is Image Then
            photo_ = picture_
        ElseIf TypeOf picture_ Is String Then
            photo_ = Image.FromFile(picture_)
        End If


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

    Public Shared Function Content(control_ As PictureBox, Optional use_image As Boolean = False) As Object
        Try
            If use_image Then
                Return PictureFromStream(control_, ".jpg", True)
            Else
                Return PictureFromStream(control_)
            End If
        Catch ex As Exception

        End Try

    End Function

    Public Shared Property casing__ As TextCase = TextCase.Capitalize
    Public Shared Function Content(control_ As Object, Optional casing_ As TextCase = TextCase.None, Optional timeValue As TimeValue = TimeValue.ShortDate) As String
        Dim casing As TextCase = casing__
        If casing_ <> Nothing Then casing = casing_

        Dim nud As NumericUpDown
        Dim h As CheckBox
        Dim d As DateTimePicker
        Dim t As TrackBar
        '		Dim html_text As HtmlInputText
        Try
            If TypeOf control_ Is NumericUpDown Then
                nud = control_
                Return nud.Value
            ElseIf TypeOf control_ Is TrackBar Then
                t = control_
                Return t.Value
            ElseIf TypeOf control_ Is CheckBox Then
                h = control_
                Return h.Checked
            ElseIf TypeOf control_ Is DateTimePicker Then
                d = control_
                Select Case timeValue
                    Case TimeValue.Day
                        Return CType(control_, DateTimePicker).Value.Day
                    Case TimeValue.Hour
                        Return CType(control_, DateTimePicker).Value.Hour
                    Case TimeValue.LongDate
                        Return CType(control_, DateTimePicker).Value.ToLongDateString
                    Case TimeValue.LongTime
                        Return CType(control_, DateTimePicker).Value.ToLongTimeString
                    Case TimeValue.Millisecond
                        Return CType(control_, DateTimePicker).Value.Millisecond
                    Case TimeValue.Minute
                        Return CType(control_, DateTimePicker).Value.Minute
                    Case TimeValue.Month
                        Return CType(control_, DateTimePicker).Value.Month
                    Case TimeValue.Second
                        Return CType(control_, DateTimePicker).Value.Second
                    Case TimeValue.ShortDate
                        Return CType(control_, DateTimePicker).Value.ToShortDateString
                    Case TimeValue.ShortTime
                        Return CType(control_, DateTimePicker).Value.ToShortTimeString
                    Case TimeValue.Year
                        Return CType(control_, DateTimePicker).Value.Year
                    Case TimeValue.DayOfWeek
                        Return CType(control_, DateTimePicker).Value.DayOfWeek
                    Case TimeValue.DayOfYear
                        Return CType(control_, DateTimePicker).Value.DayOfYear
                End Select
            ElseIf TypeOf control_ Is ComboBox Then
                Return TransformText(control_.Text, casing)
            ElseIf TypeOf control_ Is ListBox Then
                Try
                    Return TransformText(CType(control_, ListBox).SelectedItem, casing)
                Catch ex As Exception

                End Try
            ElseIf TypeOf control_ Is String Then
                Return ReadText(control_)
            Else
                Try
                    Return TransformText(control_.Text, casing)
                Catch ex As Exception
                End Try
            End If
        Catch
        End Try
    End Function

    Public Shared Sub Clear(c_ As Array)
        If c_.Length > 0 Then
            With c_
                For i = 0 To .Length - 1
                    Clear(c_(i))
                Next
            End With
        End If
    End Sub

    Public Shared Sub Clear(c_ As Object, Optional initial_string As String = "")

        Dim c As ComboBox
        Dim l As ListBox
        Dim t As TextBox
        Dim p As PictureBox
        Dim h As CheckBox
        Dim g As DataGridView

        If TypeOf c_ Is DataGridView Then
            g = c_
            Try
                g.DataSource = Nothing
            Catch
            End Try
        ElseIf TypeOf c_ Is CheckBox Then
            h = c_
            h.Checked = False
        ElseIf TypeOf c_ Is ComboBox Then
            c = c_
            'Try
            '	c.DataSource = Nothing
            'Catch ex As Exception
            'End Try
            'c.Items.Clear()
            c.Text = initial_string
        ElseIf TypeOf c_ Is ListBox Then
            l = c_
            Try
                l.DataSource = Nothing
            Catch ex As Exception
            End Try
            l.Items.Clear()
        ElseIf TypeOf c_ Is TextBox Then
            t = c_
            t.Text = initial_string
        ElseIf TypeOf c_ Is PictureBox Then
            p = c_
            Try
                p.Image = Nothing
            Catch ex As Exception
            End Try
            Try
                p.BackgroundImage = Nothing
            Catch ex As Exception
            End Try
        ElseIf TypeOf c_ Is NumericUpDown Then
            Try
                If IsNumeric(initial_string) Then '/
                    CType(c_, NumericUpDown).Value = initial_string
                Else
                    CType(c_, NumericUpDown).Value = CType(c_, NumericUpDown).Minimum
                End If
            Catch ex As Exception

            End Try
        Else
            WriteText(c_, initial_string)
        End If
    End Sub

    ''' <summary>
    ''' Populates drop with drop-text version of boolean, depending on the pattern. Same as PopulateBooleanDrop.
    ''' </summary>
    ''' <param name="d_">ComboBox to fill</param>
    ''' <param name="firstItemIsEmpty">Should 's first item be empty?</param>
    ''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
    Public Shared Sub BooleanDrop(d_ As ComboBox, Optional pattern_ As String = "always/never", Optional firstItemIsEmpty As Boolean = False)
        PopulateBooleanDrop(d_, pattern_, firstItemIsEmpty)

    End Sub

    Public Shared Sub BooleanDrop(d_ As ComboBox, Optional pattern_ As DropTextPattern = DropTextPattern.AlwaysNever, Optional firstItemIsEmpty As Boolean = False)
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
    ''' <summary>
    ''' Populates drop with drop-text version of boolean, depending on the pattern.
    ''' </summary>
    ''' <param name="d_">ComboBox to fill</param>
    ''' <param name="firstItemIsEmpty">Should 's first item be empty?</param>
    ''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
    Private Shared Sub PopulateBooleanDrop(d_ As ComboBox, Optional pattern_ As String = "always/never", Optional firstItemIsEmpty As Boolean = False)
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
    Public Sub PopulateBEDrop(d_ As ComboBox, Optional firstItemIsEmpty As Boolean = False)
        If d_.Items.Count > 0 Then Exit Sub

        With d_.Items
            Try
                If firstItemIsEmpty = True And d_.DataSource Is Nothing Then .Add("")
            Catch
            End Try
            .Add("Begin")
            .Add("End")
        End With
    End Sub


    ''' <summary>
    ''' Gets items in combobox and adds them to a list of string or array
    ''' </summary>
    ''' <param name="c"></param>
    ''' <param name="returnAs">List(Of String) OR Array</param>
    ''' <returns>List(Of String) OR Array</returns>
    Public Shared Function ComboToList(c As ComboBox, Optional returnAs As ReturnInfo = ReturnInfo.AsArray)
        Dim l As New List(Of String)
        With c
            For i = 0 To .Items.Count - 1
                l.Add(c.Items(i))
            Next
        End With
        If returnAs = ReturnInfo.AsListOfString Then
            Return l
        ElseIf returnAs = ReturnInfo.AsArray Then
            Return l.ToArray
        End If
    End Function









End Class
