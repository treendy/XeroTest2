Imports System.IO
Imports System.Net

Public Class _default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click

        Dim Response As String = OAuthWebRequest(clsOAuth.Method.GET, "https://api.xero.com/oauth/requesttoken", String.Empty)

        'Do something with this auth token response to allow access to private app

        'Get list of items and show in txtResult.text

    End Sub

    Public Function OAuthWebRequest(ByVal RequestMethod As clsOAuth.Method, ByVal url As String, ByVal PostData As String) As String
        Dim OutURL As String = String.Empty
        Dim QueryString As String = String.Empty
        Dim ReturnValue As String = String.Empty
        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidateCertificate)
        If RequestMethod = clsOAuth.Method.POST Then
            If PostData.Length > 0 Then
                Dim qs As NameValueCollection = HttpUtility.ParseQueryString(PostData)
                PostData = String.Empty
                For Each Key As String In qs.AllKeys
                    If PostData.Length > 0 Then
                        PostData &= "&"
                    End If
                    qs(Key) = HttpUtility.UrlDecode(qs(Key))
                    qs(Key) = clsOAuth.OAuthUrlEncode(qs(Key))
                    PostData &= Key + "=" + qs(Key)
                Next
                If url.IndexOf("?") > 0 Then
                    url &= "&"
                Else
                    url &= "?"
                End If
                url &= PostData
            End If
        End If

        Dim RequestUri As New Uri(url)
        Dim Nonce As String = clsOAuth.GenerateNonce()
        Dim TimeStamp As String = clsOAuth.GenerateTimeStamp()
        Dim ConsumerKey As String = "L3IDVMQ8S0IEOAGCXLFYCOO6X7TEF9"
        Dim ConsumerSecret As String = "O0TEWNI5EP8OWDAL45SX4OKXMPWJ3C"
        Dim Token As String = ""
        Dim TokenSecret As String = ""
        Dim CallbackUrl As String = ""
        Dim Verifier As String = ""
        Dim cAuth As New clsOAuth()
        Dim Sig As String = cAuth.GenerateSignature(RequestUri, ConsumerKey, ConsumerSecret, Token, TokenSecret, RequestMethod.ToString, TimeStamp, Nonce, OutURL, QueryString, CallbackUrl, Verifier)
        QueryString &= "&oauth_signature=" + clsOAuth.OAuthUrlEncode(Sig)
        If RequestMethod = clsOAuth.Method.POST Then
            PostData = QueryString
            QueryString = String.Empty
        End If

        If QueryString.Length > 0 Then
            OutURL &= "?"
        End If

        ReturnValue = WebRequestCust(RequestMethod, OutURL + QueryString, PostData)

        Return ReturnValue
    End Function

    Private Function ValidateCertificate(ByVal sender As Object, ByVal certificate As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
        Return True
    End Function

    Public Function WebRequestCust(ByVal RequestMethod As clsOAuth.Method, ByVal Url As String, ByVal PostData As String) As String

        Try
            Dim Request As HttpWebRequest = TryCast(System.Net.WebRequest.Create(Url), HttpWebRequest)
            Dim p As System.Net.WebProxy = Nothing
            Request.Method = RequestMethod.ToString()
            Request.ServicePoint.Expect100Continue = False
            'If Globals.Proxy_Username <> String.Empty And Globals.Proxy_Password <> String.Empty Then
            '    p = New System.Net.WebProxy
            '    p.Credentials = New NetworkCredential(Globals.Proxy_Username, Globals.Proxy_Password)
            '    Request.Proxy = p
            'End If

            If RequestMethod = clsOAuth.Method.POST Then
                Request.ContentType = "application/x-www-form-urlencoded"
                Using RequestWriter As New StreamWriter(Request.GetRequestStream())
                    RequestWriter.Write(PostData)
                End Using
            End If

            Dim wr As System.Net.WebResponse = Request.GetResponse
            'Globals.API_HourlyLimit = wr.Headers("X-RateLimit-Limit")
            'Globals.API_RemainingHits = wr.Headers("X-RateLimit-Remaining")
            'Globals.API_Reset = wr.Headers("X-RateLimit-Reset")

            Using ResponseReader As New StreamReader(wr.GetResponseStream())
                Return ResponseReader.ReadToEnd
            End Using

        Catch ex As Exception
            Dim Message As String = Nothing
            If TypeOf ex Is WebException Then
                Try
                    Dim Doc As New System.Xml.XmlDocument
                    Doc.LoadXml(New StreamReader(CType(ex, WebException).Response.GetResponseStream, Encoding.UTF8).ReadToEnd)
                    Message = Doc.SelectSingleNode("hash/error").InnerText
                Catch
                End Try
            End If
            'Dim tax As New TwitterAPIException(Message, ex)
            'With tax
            '    .Url = Url
            '    .Method = RequestMethod.ToString
            '    .AuthType = "OAUTH"
            '    .Response = Nothing
            '    .Status = Nothing
            'End With
            'Throw tax
        End Try

    End Function

End Class