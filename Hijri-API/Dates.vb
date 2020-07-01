Imports System
Imports System.Web
Imports System.Diagnostics
Imports System.Globalization
Imports System.Data
Imports System.Collections


Namespace Converter
    ''' <summary>
    ''' Summary description for Dates.
    ''' </summary>

    Public Class Dates
        Private cur As HttpContext

        Private Const startGreg As Integer = 1900
        Private Const endGreg As Integer = 2100
        Private allFormats As String() = {"yyyy/MM/dd", "yyyy/M/d", "dd/MM/yyyy", "d/M/yyyy", "dd/M/yyyy", "d/MM/yyyy", _
         "yyyy-MM-dd", "yyyy-M-d", "dd-MM-yyyy", "d-M-yyyy", "dd-M-yyyy", "d-MM-yyyy", _
         "yyyy MM dd", "yyyy M d", "dd MM yyyy", "d M yyyy", "dd M yyyy", "d MM yyyy"}
        Private arCul As CultureInfo
        Private enCul As CultureInfo
        Private h As UmAlQuraCalendar
        Private g As GregorianCalendar

        Public Sub New()
            cur = HttpContext.Current

            arCul = New CultureInfo("ar-SA")
            enCul = New CultureInfo("en-US")

            h = New UmAlQuraCalendar()
            g = New GregorianCalendar(GregorianCalendarTypes.USEnglish)


            arCul.DateTimeFormat.Calendar = h
        End Sub

        ''' <summary>
        ''' Check if string is hijri date and then return true 
        ''' </summary>
        ''' <param name="hijri"></param>
        ''' <returns></returns>
        Public Function IsHijri(ByVal hijri As String) As Boolean
            If hijri.Length <= 0 Then

                cur.Trace.Warn("IsHijri Error: Date String is Empty")
                Return False
            End If
            Try
                Dim tempDate As DateTime = DateTime.ParseExact(hijri, allFormats, arCul.DateTimeFormat, DateTimeStyles.AllowWhiteSpaces)
                If tempDate.Year >= startGreg AndAlso tempDate.Year <= endGreg Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                cur.Trace.Warn("IsHijri Error :" + hijri.ToString() + vbLf + ex.Message)
                Return False
            End Try

        End Function
        ''' <summary>
        ''' Check if string is Gregorian date and then return true 
        ''' </summary>
        ''' <param name="greg"></param>
        ''' <returns></returns>
        Public Function IsGreg(ByVal greg As String) As Boolean
            If greg.Length <= 0 Then

                cur.Trace.Warn("IsGreg :Date String is Empty")
                Return False
            End If
            Try
                Dim tempDate As DateTime = DateTime.ParseExact(greg, allFormats, enCul.DateTimeFormat, DateTimeStyles.AllowWhiteSpaces)
                If tempDate.Year >= startGreg AndAlso tempDate.Year <= endGreg Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                cur.Trace.Warn("IsGreg Error :" + greg.ToString() + vbLf + ex.Message)
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Convert Hijri Date to it's equivalent Gregorian Date
        ''' </summary>
        ''' <param name="hijri"></param>
        ''' <returns></returns>
        Public Function HijriToGreg(ByVal hijri As String) As String

            If hijri.Length <= 0 Then

                cur.Trace.Warn("HijriToGreg :Date String is Empty")
                Return ""
            End If
            Try
                Dim tempDate As DateTime = DateTime.ParseExact(hijri, allFormats, arCul.DateTimeFormat, DateTimeStyles.AllowWhiteSpaces)
                Return tempDate.ToString("yyyy/MM/dd", enCul.DateTimeFormat)
            Catch ex As Exception
                cur.Trace.Warn("HijriToGreg :" + hijri.ToString() + vbLf + ex.Message)
                Return ""
            End Try
        End Function
        ''' <summary>
        ''' Convert Hijri Date to it's equivalent Gregorian Date and return it in specified format
        ''' </summary>
        ''' <param name="hijri"></param>
        ''' <param name="format"></param>
        ''' <returns></returns>
        Public Function HijriToGreg(ByVal hijri As String, ByVal format As String) As String

            If hijri.Length <= 0 Then

                cur.Trace.Warn("HijriToGreg :Date String is Empty")
                Return ""
            End If
            Try
                Dim tempDate As DateTime = DateTime.ParseExact(hijri, allFormats, arCul.DateTimeFormat, DateTimeStyles.AllowWhiteSpaces)

                Return tempDate.ToString(format, enCul.DateTimeFormat)
            Catch ex As Exception
                cur.Trace.Warn("HijriToGreg :" + hijri.ToString() + vbLf + ex.Message)
                Return ""
            End Try
        End Function
        ''' <summary>
        ''' Convert Gregoian Date to it's equivalent Hijir Date
        ''' </summary>
        ''' <param name="greg"></param>
        ''' <returns></returns>
        Public Function GregToHijri(ByVal greg As String) As String

            If greg.Length <= 0 Then

                cur.Trace.Warn("GregToHijri :Date String is Empty")
                Return ""
            End If
            Try
                Dim tempDate As DateTime = DateTime.ParseExact(greg, allFormats, enCul.DateTimeFormat, DateTimeStyles.AllowWhiteSpaces)

                Return tempDate.ToString("yyyy/MM/dd", arCul.DateTimeFormat)
            Catch ex As Exception
                cur.Trace.Warn("GregToHijri :" + greg.ToString() + vbLf + ex.Message)
                Return ""
            End Try
        End Function
        Public Function WeekDay(ByVal d1 As String) As String
            Dim date1 As DateTime
            Try
                date1 = DateTime.ParseExact(d1, allFormats, enCul.DateTimeFormat, DateTimeStyles.AllowWhiteSpaces)
                Return date1.DayOfWeek.ToString
            Catch ex As Exception
                Return "This is not a valid Hijri date"
            End Try


        End Function

        ''' <summary>
        ''' Convert Hijri Date to it's equivalent Gregorian Date and return it in specified format
        ''' </summary>
        ''' <param name="greg"></param>
        ''' <param name="format"></param>
        ''' <returns></returns>
        Public Function GregToHijri(ByVal greg As String, ByVal format As String) As String

            If greg.Length <= 0 Then

                cur.Trace.Warn("GregToHijri :Date String is Empty")
                Return ""
            End If
            Try

                Dim tempDate As DateTime = DateTime.ParseExact(greg, allFormats, enCul.DateTimeFormat, DateTimeStyles.AllowWhiteSpaces)

                Return tempDate.ToString(format, arCul.DateTimeFormat)
            Catch ex As Exception
                cur.Trace.Warn("GregToHijri :" + greg.ToString() + vbLf + ex.Message)
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Convert Hijri Date to it's equivalent Gegorian Date and return it in specified format according to Zurich
        ''' </summary>
        ''' <param name="d"></param>
        ''' <param name="m"></param>
        ''' <param name="y"></param>
        ''' <returns></returns>
        Public Function H2GZurich(ByVal d As Integer, ByVal m As Integer, ByVal y As Long) As String
            Dim jd As Long, WDay As String, result As String, l As Long, n As Long, j As Long, i As Long, k As Long
            jd = intPart((11 * y + 3) / 30) + 354 * y + 30 * m - intPart((m - 1) / 2) + d + 1948440 - 385
            WDay = ZWeekday(jd Mod 7)
            If (jd > 2299160) Then
                l = jd + 68569
                n = intPart((4 * l) / 146097)
                l = l - intPart((146097 * n + 3) / 4)
                i = intPart((4000 * (l + 1)) / 1461001)
                l = l - intPart((1461 * i) / 4) + 31
                j = intPart((80 * l) / 2447)
                d = l - intPart((2447 * j) / 80)
                l = intPart(j / 11)
                m = j + 2 - 12 * l
                y = 100 * (n - 49) + i + l
            Else
                j = jd + 1402
                k = intPart((j - 1) / 1461)
                l = j - 1461 * k
                n = intPart((l - 1) / 365) - intPart(l / 1461)
                i = l - 365 * n + 30
                j = intPart((80 * i) / 2447)
                d = i - intPart((2447 * j) / 80)
                i = intPart(j / 11)
                m = j + 2 - 12 * i
                y = 4 * k + n + i - 4716
            End If

            Dim dday As String, mmonth As String, yyear As String
            dday = CStr(d)
            If dday.Length = 1 Then dday = "0" + dday
            mmonth = CStr(m)
            If mmonth.Length = 1 Then mmonth = "0" + mmonth
            yyear = CStr(y)
            Do Until yyear.Length = 4
                yyear = "0" + yyear
            Loop
            result = dday + mmonth + yyear + WDay
            Return result
        End Function
        ''' <summary>
        ''' Convert Gregorian Date to it's equivalent Hijri Date and return it in specified format according to Zurich
        ''' </summary>
        ''' <param name="d"></param>
        ''' <param name="m"></param>
        ''' <param name="y"></param>
        ''' <returns></returns>
        Public Function G2HZurich(ByVal d As Integer, ByVal m As Integer, ByVal y As Long) As String
            Dim jd As Long, WDay As String, result As String, l As Long, n As Long, j As Long
            If ((y > 1582) Or ((y = 1582) And (m > 10)) Or ((y = 1582) And (m = 10) And (d > 14))) Then
                jd = intPart((1461 * (y + 4800 + intPart((m - 14) / 12))) / 4) + intPart((367 * (m - 2 - 12 * (intPart((m - 14) / 12)))) / 12) - intPart((3 * (intPart((y + 4900 + intPart((m - 14) / 12)) / 100))) / 4) + d - 32075
            Else
                jd = 367 * y - intPart((7 * (y + 5001 + intPart((m - 9) / 7))) / 4) + intPart((275 * m) / 9) + d + 1729777
            End If
            WDay = ZWeekday(jd Mod 7)
            l = jd - 1948440 + 10632
            n = intPart((l - 1) / 10631)
            l = l - 10631 * n + 354
            j = (intPart((10985 - l) / 5316)) * (intPart((50 * l) / 17719)) + (intPart(l / 5670)) * (intPart((43 * l) / 15238))
            l = l - (intPart((30 - j) / 15)) * (intPart((17719 * j) / 50)) - (intPart(j / 16)) * (intPart((15238 * j) / 43)) + 29
            m = intPart((24 * l) / 709)
            d = l - intPart((709 * m) / 24)
            y = 30 * n + j - 30
            Dim dday As String, mmonth As String, yyear As String
            dday = CStr(d)
            If dday.Length = 1 Then dday = "0" + dday
            mmonth = CStr(m)
            If mmonth.Length = 1 Then mmonth = "0" + mmonth
            yyear = CStr(y)
            Do Until yyear.Length = 4
                yyear = "0" + yyear
            Loop
            result = dday + mmonth + yyear + WDay
            Return result
        End Function
        ''' <summary>
        ''' Flooring
        ''' </summary>
        ''' <param name="floatNum"></param>
        ''' <returns></returns>
        Public Function intPart(ByVal floatNum As Double)
            If (floatNum < -0.0000001) Then
                Return Math.Ceiling(floatNum - 0.0000001)
            Else
                Return Math.Floor(floatNum + 0.0000001)
            End If
        End Function
        ''' <summary>
        ''' Weekday
        ''' </summary>
        ''' <param name="x"></param>
        ''' <returns></returns>
        Public Function ZWeekday(ByVal x As Long) As String
            Dim y As String = ""
            If x = 0 Then y = "Monday"
            If x = 1 Then y = "Tuesday"
            If x = 2 Then y = "Wednesday"
            If x = 3 Then y = "Thursday"
            If x = 4 Then y = "Friday"
            If x = 5 Then y = "Saturday"
            If x = 6 Then y = "Sunday"
            Return y
        End Function

    End Class
End Namespace