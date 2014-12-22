Imports SeasideResearch.LibCurlNet
Imports System.Text.RegularExpressions
Imports System.Web

Public Class clsHttpClient

    'by Michael Evanchik so you can use HTTP with Tor
    'to avoid memory leaks, try to _ALWAYS_ .dispose of this POS!

    Implements IDisposable
    Protected m_bDisposed As Boolean

    'make sure to add these to your prj!!!
    'Curl.GlobalInit()
    'Curl.GlobalCleanup()

    Public m_cEasy As Easy
    'Public m_sUserAgent As String = "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.0.4) Gecko/20060508 Firefox/1.5.0.4"
    Public m_sUserAgent As String = "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.0.12) Gecko/20070508 Firefox/1.5.0.12"
    Public m_sPageData As String
    Public m_sCookieFile As String
    Public m_bDebug As Boolean = False
    Private m_mpf As New MultiPartForm
    Private m_bTor As Boolean = False

    Public Property RouteWithTor() As Boolean
        Get
            RouteWithTor = m_bTor
        End Get
        Set(ByVal usetor As Boolean)
            m_bTor = Not m_bTor
            If m_bTor Then
                m_cEasy.SetOpt(CURLoption.CURLOPT_PROXY, "localhost:9050")
                m_cEasy.SetOpt(CURLoption.CURLOPT_PROXYTYPE, CURLproxyType.CURLPROXY_SOCKS5)
            Else
                m_cEasy.SetOpt(CURLoption.CURLOPT_PROXY, vbNullString)
            End If
        End Set
    End Property

    Protected Overridable Sub Dispose(ByVal disposing As Boolean)

        If disposing Then
            ' Call dispose on any objects referenced by this object
        End If
        ' Release unmanaged resources
        m_cEasy.Cleanup()

    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        m_bDisposed = True
        ' Take off finalization queue
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        Me.Dispose(False)
    End Sub

    Public Sub New()

        m_cEasy = New Easy

        With m_cEasy
            .SetOpt(CURLoption.CURLOPT_TIMEOUT, 180)
            .SetOpt(CURLoption.CURLOPT_FOLLOWLOCATION, 1)
            .SetOpt(CURLoption.CURLOPT_USERAGENT, m_sUserAgent)
        End With

    End Sub

    Public Sub NewMultiPartForm()

        m_mpf.Free()
        m_mpf = New MultiPartForm

    End Sub

    Public Sub AddMPFSection(ByVal ParamArray args() As Object)

        m_mpf.AddSection(args)

    End Sub

    Private Function OnWriteData(ByVal buf() As Byte, ByVal size As Int32, ByVal nmemb As Int32, ByVal extraData As Object) As Int32

        m_sPageData &= System.Text.Encoding.Default.GetString(buf)
        Return size * nmemb

    End Function

    Public Function OnHeaderData(ByVal buf() As Byte, ByVal size As Int32, ByVal nmemb As Int32, ByVal extraData As Object) As Int32

        If m_bDebug Then Debug.Write("[?] " & System.Text.Encoding.Default.GetString(buf))
        Return size * nmemb

    End Function
    'should do OnHeaderData and OnDebugData, but i'm too lazy for all that

    Public Sub DumpCookies(Optional ByVal delete As Boolean = False)

        'force dump of cookies to file and reinit curl handle
        m_cEasy.Cleanup()
        If delete And System.IO.File.Exists(m_sCookieFile) Then Kill(m_sCookieFile)
        m_cEasy = New Easy
        With m_cEasy
            .SetOpt(CURLoption.CURLOPT_TIMEOUT, 180)
            .SetOpt(CURLoption.CURLOPT_WRITEFUNCTION, New Easy.WriteFunction(AddressOf OnWriteData))
            .SetOpt(CURLoption.CURLOPT_HEADERFUNCTION, New Easy.HeaderFunction(AddressOf OnHeaderData))
            .SetOpt(CURLoption.CURLOPT_FOLLOWLOCATION, 1)
            .SetOpt(CURLoption.CURLOPT_USERAGENT, m_sUserAgent)
        End With

    End Sub

    Public Sub EditCookie(ByVal name As String, ByVal data As String)

        Dim fn As Integer = FreeFile()
        Dim line As String
        Dim temp As String
        Dim flag As Boolean = False
        Dim lines As New Collection

        DumpCookies()

        'read cookies to mem, and edit the wanted one
        Try
            FileOpen(fn, m_sCookieFile, OpenMode.Input)
            Do While Not EOF(fn)
                line = LineInput(fn)
                If (Left(line, 1) <> "#") And (Len(Trim(line)) > 0) Then
                    If name = line.Split(vbTab)(5) Then
                        temp = line.Split(vbTab)(0)
                        temp &= vbTab & line.Split(vbTab)(1)
                        temp &= vbTab & line.Split(vbTab)(2)
                        temp &= vbTab & line.Split(vbTab)(3)
                        temp &= vbTab & line.Split(vbTab)(4)
                        temp &= vbTab & line.Split(vbTab)(5)
                        temp &= vbTab & data
                        line = temp
                        flag = True
                    End If
                    lines.Add(line)
                End If
            Loop
            FileClose(fn)
        Catch ex As Exception
            Debug.Print("[!] error reading cookie file")
        End Try

        If flag = False Then lines.Add(vbTab & vbTab & vbTab & vbTab & vbTab & name & data)

        'rewrite edited file
        FileOpen(fn, m_sCookieFile, OpenMode.Output)
        For Each line In lines
            PrintLine(fn, line)
        Next
        FileClose(fn)

    End Sub

    Private Sub Perform(Optional ByVal status As String = "")

        m_sPageData = vbNullString
        Debug.WriteLine("[*] http.client: " & status)
        With m_cEasy
            .SetOpt(CURLoption.CURLOPT_FILETIME, True)
            .SetOpt(CURLoption.CURLOPT_CAINFO, "ca-bundle.crt")
            .SetOpt(CURLoption.CURLOPT_WRITEFUNCTION, New Easy.WriteFunction(AddressOf OnWriteData))
            .SetOpt(CURLoption.CURLOPT_HEADERFUNCTION, New Easy.HeaderFunction(AddressOf OnHeaderData))
            .SetOpt(CURLoption.CURLOPT_COOKIEFILE, m_sCookieFile)
            .SetOpt(CURLoption.CURLOPT_COOKIEJAR, m_sCookieFile)
        End With


        m_cEasy.Perform()

    End Sub

    Public Function Fetch(ByRef doc As String, ByVal url As String, Optional ByVal ref As String = vbNullString, Optional ByVal txt As String = vbNullString) As Boolean

        If Len(txt) = 0 Then txt = "fetching url: " & url
        With m_cEasy
            .SetOpt(CURLoption.CURLOPT_POST, 0)
            .SetOpt(CURLoption.CURLOPT_URL, url)
            If Len(ref) > 0 Then _
                .SetOpt(CURLoption.CURLOPT_REFERER, ref)
        End With
        Perform(txt)
        doc = m_sPageData
        Fetch = CBool(Len(doc) > 0)

    End Function

    Public Function re_Fetch(ByRef doc As String, ByVal pattern As String, ByVal url As String, Optional ByVal ref As String = vbNullString, Optional ByVal options As System.Text.RegularExpressions.RegexOptions = RegexOptions.IgnoreCase, Optional ByVal txt As String = vbNullString) As MatchCollection

        Dim re As New Regex(pattern, options)
        Dim mc As MatchCollection

        mc = Nothing
        Try
            If Len(txt) = 0 Then txt = "fetching url: " & url
            With m_cEasy
                .SetOpt(CURLoption.CURLOPT_POST, 0)
                .SetOpt(CURLoption.CURLOPT_URL, url)
                If Len(ref) > 0 Then _
                    .SetOpt(CURLoption.CURLOPT_REFERER, ref)
            End With
            Perform(txt)
            doc = m_sPageData
            mc = re.Matches(doc)
        Catch ex As Exception
            Debug.Print("[!] re_fetch : " & ex.Message.ToString.ToLower)
        Finally
            re_Fetch = mc
        End Try

    End Function

    Public Function GetInfo(ByVal info As SeasideResearch.LibCurlNet.CURLINFO)

        Dim ret As Object = Nothing
        m_cEasy.GetInfo(info, ret)
        GetInfo = ret

    End Function

    Public Function PostMPFData(ByRef doc As String, ByVal url As String, Optional ByVal ref As String = vbNullString, Optional ByVal txt As String = vbNullString) As Boolean

        If Len(txt) = 0 Then txt = "posting multipart form data to url: " & url
        With m_cEasy
            .SetOpt(CURLoption.CURLOPT_HTTPPOST, m_mpf)
            .SetOpt(CURLoption.CURLOPT_URL, url)
            If Len(ref) > 0 Then _
                .SetOpt(CURLoption.CURLOPT_REFERER, ref)
        End With
        Perform(txt)
        doc = m_sPageData
        PostMPFData = (Len(doc) > 0)

    End Function

    Public Function PostData(ByRef doc As String, ByVal url As String, Optional ByVal ref As String = vbNullString, Optional ByVal post As String = vbNullString, Optional ByVal txt As String = vbNullString) As Boolean

        If Len(txt) = 0 Then txt = "posting to url: " & url
        With m_cEasy
            .SetOpt(CURLoption.CURLOPT_POST, 1)
            .SetOpt(CURLoption.CURLOPT_POSTFIELDS, post)
            .SetOpt(CURLoption.CURLOPT_URL, url)
            If Len(ref) > 0 Then _
                .SetOpt(CURLoption.CURLOPT_REFERER, ref)
        End With
        Perform(txt)
        doc = m_sPageData
        PostData = (Len(doc) > 0)

    End Function

    Public Function re_PostData(ByRef doc As String, ByVal pattern As String, ByVal url As String, Optional ByVal ref As String = vbNullString, Optional ByVal post As String = vbNullString, Optional ByVal options As System.Text.RegularExpressions.RegexOptions = RegexOptions.IgnoreCase, Optional ByVal txt As String = vbNullString) As MatchCollection

        Dim re As New Regex(pattern, options)
        Dim mc As MatchCollection = Nothing

        If Len(txt) = 0 Then txt = "posting to url: " & url
        With m_cEasy
            .SetOpt(CURLoption.CURLOPT_POST, 1)
            .SetOpt(CURLoption.CURLOPT_POSTFIELDS, post)
            .SetOpt(CURLoption.CURLOPT_URL, url)
            If Len(ref) > 0 Then _
                .SetOpt(CURLoption.CURLOPT_REFERER, ref)
        End With
        Perform(txt)
        doc = m_sPageData
        If Len(doc) Then mc = re.Matches(doc)
        re_PostData = mc

    End Function

    Public Sub SaveAndRun(ByVal filename As String, ByVal runit As Boolean)

        Dim fn As Integer = FreeFile()

        FileOpen(fn, filename, OpenMode.Output)
        Print(fn, m_sPageData)
        FileClose(fn)

        If runit Then System.Diagnostics.Process.Start(filename)

    End Sub

    Public Function UrlDecode(ByVal txt As String) As String

        UrlDecode = HttpUtility.UrlDecode(txt)

    End Function

    Public Function UrlEncode(ByVal txt As String) As String

        UrlEncode = HttpUtility.UrlEncode(txt)

    End Function

End Class
