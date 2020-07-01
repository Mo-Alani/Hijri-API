Imports Hijri_API.Converter
Imports Hijri_API.Converter.Dates
Imports System.Net

Public Class WebForm1
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
    Dim dd As New Dates

    Private Sub Page_LoadComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LoadComplete
        Dim d As Integer, m As Integer, y As Integer, conv As String, result As String = "Error in Type", yy As String = "", Auth As String = ""
        d = CInt(Request.QueryString.Item("day"))
        m = CInt(Request.QueryString.Item("month"))
        y = CInt(Request.QueryString.Item("year"))
        conv = Request.QueryString.Item("type")
        yy = CStr(y)
        Auth = Request.QueryString.Item("auth")
        If Auth = "" Or d = 0 Or m = 0 Or y = 0 Or conv = "" Then
            result = "Parameter Error"
            Auth = "Error"
            GoTo ErrorJump
        End If


        Do Until yy.Length = 4
            yy = "0" + yy
        Loop
        If conv IsNot Nothing Then conv = conv.ToLower
        If conv = "h2g" Then
            If y > 1355 And y < 1501 Then
                result = dd.HijriToGreg(CStr(d) + "-" + CStr(m) + "-" + yy, "ddMMyyyy")
                result = result + dd.WeekDay(result)
            Else
                result = dd.H2GZurich(d, m, y)
            End If
        End If
        If conv = "g2h" Then
            If (y < 2077 And y > 1937) Or (y = 1937 And m > 3) Or (y = 1937 And m = 3 And d > 13) Or (y = 2077 And m < 11) Or (y = 2077 And m = 11 And d < 17) Then
                result = dd.GregToHijri(CStr(d) + "-" + CStr(m) + "-" + yy, "ddMMyyyy") + dd.WeekDay(CStr(d) + "-" + CStr(m) + "-" + CStr(y))

            Else
                result = dd.G2HZurich(d, m, y)
            End If

        End If
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