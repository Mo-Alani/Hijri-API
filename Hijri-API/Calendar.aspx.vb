Imports Hijri_API.Converter
Imports Hijri_API.Converter.Dates
Imports System.Net

Public Class Calendar
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
    Dim dd As New Dates
    Private Sub Calendar_LoadComplete(sender As Object, e As EventArgs) Handles Me.LoadComplete
        Dim d As Integer, m As Integer, y As Integer, res As String, result As String = "Error in Type", yy As String = "", Auth As String = ""
        d = 1
        m = CInt(Request.QueryString.Item("month"))
        y = CInt(Request.QueryString.Item("year"))

        yy = CStr(y)
        Auth = Request.QueryString.Item("auth")
        If Auth = "" Or m = 0 Or m > 12 Or y = 0 Then
            result = "Parameter Error"
            Auth = "Error"
            GoTo ErrorJump
        End If


        Do Until yy.Length = 4
            yy = "0" + yy
        Loop
        result = ""
        For i = 1 To 30
            If y > 1355 And y < 1501 Then
                res = dd.HijriToGreg(CStr(i) + "-" + CStr(m) + "-" + yy, "ddMMyyyy")
                If i = 1 Then res = res + dd.WeekDay(res)
            Else
                res = dd.H2GZurich(i, m, y)
                If i <> 1 Then res = res.Substring(0, 8)
            End If
            'If m < 12 Then
            'm = m + 1
            'Else
            '           m = 1
            '          y = y + 1

            '         End If
            'yy = CStr(y)
            'Do Until yy.Length = 4
            'yy = "0" + yy
            'Loop
            'If y > 1355 And y < 1501 Then
            '  res = dd.HijriToGreg(CStr(d) + "-" + CStr(m) + "-" + yy, "ddMMyyyy")
            '   res = res + dd.WeekDay(res)
            'Else
            '   res = dd.H2GZurich(d, m, y)
            'End If

            If result <> "" Then result = result + "-"
            result = result + res
        Next i

ErrorJump:
        OutputLabel.Text = result
        Dim wb As New WebClient
        AddHandler wb.DownloadStringCompleted, AddressOf wb_Completed
        Dim tempuri As String = ""
        tempuri = "http://analytics.hijrical.org/piwik.php?idsite=2&url=http://www.hijri-api.info/&token_auth=15e93a345670b2b2d6b5043c1faf9463&rec=1&cip=" + Request.ServerVariables("REMOTE_ADDR") + "&action_name=" + Auth
        wb.DownloadStringAsync(New Uri(tempuri))
    End Sub

    Private Sub wb_Completed(sender As Object, e As DownloadStringCompletedEventArgs)

    End Sub


End Class