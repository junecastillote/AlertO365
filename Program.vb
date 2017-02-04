Imports System.Collections.Generic
Imports System.Collections
Imports System.Configuration
Imports System.Text
Imports System.Net
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Security
Imports Microsoft.Exchange.ServiceStatus.TenantCommunications.Data
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Json
Imports System.Data
Imports System.Net.Mail

Class Program
#Region "VARIABLES"
    Shared SendViaEmail As String = ""
    Shared serviceUrl As String = ""
    Shared domainName As String = ""
    Shared userId As String = ""
    Shared password As String = ""
    Shared cookie As String = ""
    Shared isAOBO As Boolean = False
    Shared pastDays As Integer = 0
    Shared SMTPServer As String = ""
    Shared SenderAddress As String = ""
    Shared RecipientAddress As String = ""
    Shared MailSubject As String = ""
    Shared Company As String = ""
    Shared str As String = ""
    Shared str2 As String = ""
    Shared reportTime As Date = Now
    Private Const RegisterMethod As String = "Register"
    Private Const RegistrationCookieProperty As String = """RegistrationCookie"":"
    Private Const GetServiceInfoForTenantDomains As String = "GetServiceInformationForTenantDomains"
    Private Const GetEventsForTenantDomainsMethod As String = "GetEventsForTenantDomains"
    Private Const RegisterRequestData As String = """userName"":""{0}"",""password"":""{1}"""
    Private Const GetServiceInfoForTenantDomainsRequestData As String = """lastCookie"":""{0}"",""tenantDomains"":[{1}],""locale"":""{2}"""
    Private Const GetEventsForTenantDomainsRequestData As String = """lastCookie"":""{0}"",""preferredEventTypes"":[{1}],""tenantDomains"":[{2}],""locale"":""{3}"",""pastDays"":""{4}"""
    Private Const GetServiceInfoMethod As String = "GetServiceInformation"
    Private Const GetEventsMethod As String = "GetEvents"
    Private Const GetServiceInfoRequestData As String = """lastCookie"":""{0}"",""locale"":""{1}"""
    Private Const GetEventsRequestData As String = """lastCookie"":""{0}"",""preferredEventTypes"":[{1}],""locale"":""{2}"",""pastDays"":""{3}"""
#End Region

    Friend Shared Sub Main(args As String())
        Try

            Console.WriteLine("===============================")
            Console.WriteLine("AlertO365 v." & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor)
            Console.WriteLine("")
            Console.WriteLine("june.castillote@gmail.com, 2015")
            Console.WriteLine("===============================")

            Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & ": Started gathering data...")


            Dim CSS_String As String = <a><![CDATA[
            <style type="text/css">
            table.titletable {font-size:18px;font-family:Verdana;color:#333333;width:100%;}
            table.tftable {font-size:12px;font-family:Verdana;color:#333333;width:100%;border-width: 1px;border-color: #729ea5;border-collapse: collapse;}
            table.tftable th {font-size:12px;font-family:Verdana;background-color:#acc8cc;border-width: 1px;padding: 8px;border-style: solid;border-color: #729ea5;text-align:middle;}
            table.tftable td {font-size:12px;font-family:Verdana;border-width: 1px;padding: 8px;border-style: solid;border-color: #729ea5;vertical-align: top}
			table.tftable td.bad {background-color:#FF9900;font-weight:bold;font-size:12px;font-family:Verdana;border-width: 1px;padding: 8px;border-style: solid;border-color: #729ea5;vertical-align: top}
			table.tftable td.good {background-color:#00FF00;font-size:12px;font-family:Verdana;border-width: 1px;padding: 8px;border-style: solid;border-color: #729ea5;vertical-align: top}
            </style>
            ]]></a>.Value

            Dim htmlFile As String = My.Application.Info.DirectoryPath & "\report\Report_" & Format(reportTime, "yyyy_MMMM_dd") & ".html"
            Dim outFile As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(htmlFile, False)
            Dim section As AppConfig = TryCast(ConfigurationManager.GetSection("ConfigSettings"), AppConfig)
            userId = section.UserName
            password = section.Password
            serviceUrl = section.ServiceURL
            domainName = section.DomainNames
            isAOBO = (section.IsAOBO = "1")
            Int32.TryParse(section.PastDays, pastDays)
            Dim domainList As String() = {domainName}
            SMTPServer = section.SMTPServer
            SenderAddress = section.SenderAddress
            RecipientAddress = section.RecipientAddress
            MailSubject = section.MailSubject
            Company = section.Company
            SendViaEmail = section.SendViaEmail

            str = str & vbNewLine & "<html> <head>" & vbNewLine & _
            "<title>" & Company & " " & MailSubject & "</title>" & vbNewLine & _
            "<meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />" & vbNewLine
            str = str & vbNewLine & CSS_String

            str = str & vbNewLine & "<table class=""titletable"" border=""0""><tr><td><img src=cid:Logo1></img><img src=cid:Logo2></img></td><td>" & Company & " " & MailSubject & " - " & reportTime & "<br>(Events Recorded in the Last " & pastDays & " Days)</td></tr></table><br>"

            ' Register
            Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & ": Set Online Registration...")
            Dim registrationInfo As RegistrationInfo = RegisterForSHDAccess(userId, password)
            cookie = registrationInfo.RegistrationCookie

            ' Get Event Data
            Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & ": Retrieving Events for the last " & pastDays & " days...")
            Dim eventInfo As EventInfo = GetEvents()
            If eventInfo IsNot Nothing Then
                Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & ": Found " & eventInfo.Events.Length & " events...")
                str = str & vbNewLine & "<table class=""tftable"" border=""1""><tr><th></th><th>Service Name</th><th>Event ID</th><th>Record Type</th><th>Status</th><th>Start Time</th><th>End Time</th><th>Last Message</th></tr>"
                For Each ev As [Event] In eventInfo.Events

                    Dim recordType As String = If((TypeOf ev Is PlannedMaintenance), "Planned Maintenance", "Service Incident")
                    Dim mySL As New SortedList()


                    Dim svcName As String
                    If ev.Id Like "EX*" Then
                        svcName = "Exchange Online"
                    ElseIf ev.Id Like "MO*" Then
                        svcName = "Office 365 Portal"
                    ElseIf ev.Id Like "LY*" Then
                        svcName = "Skype for Business"
                    ElseIf ev.Id Like "SP*" Then
                        svcName = "Sharepoint Online"
                    ElseIf ev.Id Like "IT*" Then
                        svcName = "InTune"
                    ElseIf ev.Id Like "PL*" Then
                        svcName = "Planner"
                    ElseIf ev.Id Like "FO*" Then
                        svcName = "Exchange Online Protection"
                    Else
                        svcName = "Undefined"
                    End If

                    str = str & "<tr>"
                    If ev.Status Like "Service restored" Then
                        str = str & "<td class = ""good""></td>"
                    Else
                        str = str & "<td class = ""bad""></td>"
                    End If

                    'str = str & "<td>" & svcName & "</td><td>" & ev.Id & "</td><td>" & recordType & "</td><td>" & ev.Status & "</td><td>" & ev.StartTime & "</td><td>" & ev.EndTime & "</td>"
                    str = str & "<td>" & svcName & "</td><td>" & ev.Id & "</td><td>" & recordType & "</td><td>" & ev.Status & "</td><td>" & Format(ev.StartTime, "MMMM dd, yyyy hh:mm:ss tt") & "</td><td>" & Format(ev.EndTime, "MMMM dd, yyyy hh:mm:ss tt") & "</td>"
                    Dim x As String
                    If ev.Messages IsNot Nothing AndAlso ev.Messages.Length > 0 Then
                        For Each msg As Message In ev.Messages
                            mySL.Add(msg.PublishedTime, msg.MessageText)
                        Next

                        ' Get a collection of the keys.
                        Dim key As ICollection = mySL.Keys
                        Dim k As DateTime
                        Dim i As Integer = 0
                        For Each k In key
                            If i = mySL.Keys.Count - 1 Then
                                x = Replace(mySL(k).ToString, vbLf, "<br>")
                                str = str & "<td>Published: " & Format(k, "MMMM dd, yyyy hh:mm:ss tt") & "<p>" & x & "</p></td></tr>"
                            End If
                            i = i + 1
                        Next k
                    End If

                    Console.WriteLine()
                Next
                str = str & vbNewLine & "</table>"
                str = str & vbNewLine & "<p><font size=""2"" face=""Tahoma""><a href=http://shaking-off-the-cobwebs.blogspot.com/2015/06/office-365-service-health-check.html>AlertO365</a>"
                Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & ": Saving report to file...")
                str2 = str.Replace("cid:Logo1", """images\Logo1.png""")
                str2 = str2.Replace("cid:Logo2", """images\Logo2.png""")
                outFile.WriteLine(str2)
                outFile.Close()
                If SendViaEmail = "Yes" Then
                    Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & ": Sending Report...")
                    SendReport()
                End If
                
            Else
                str = str & vbNewLine & "<table class=""tftable"" border=""1""><tr><th>No events found</th></tr></table>"
                Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & ": No Events Found...")
                Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & ": Saving report to file...")
                str2 = str.Replace("cid:Logo1", """images\Logo1.png""")
                str2 = str2.Replace("cid:Logo2", """images\Logo2.png""")
                outFile.WriteLine(str2)
                outFile.Close()

                If SendViaEmail = "Yes" Then
                    Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & ": Sending Report...")
                    SendReport()
                End If
            End If

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Shared Sub SendReport()
        Dim Logo1 As New LinkedResource(My.Application.Info.DirectoryPath & "\report\images\Logo1.png")
        Logo1.ContentId = "Logo1"
        Dim Logo2 As New LinkedResource(My.Application.Info.DirectoryPath & "\report\images\Logo2.png")
        Logo2.ContentId = "Logo2"

        Dim recipientsList As String() = Split(RecipientAddress, ",")

        Dim htmlView As AlternateView
        Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString("Cannot view this message as plain text", Nothing, "text/plain")

        Try
            'Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & " Sending Report via Email..")
            Dim SmtpServerClient As New SmtpClient()
            Dim mail As New MailMessage()

            SmtpServerClient.Host = SMTPServer

            htmlView = AlternateView.CreateAlternateViewFromString(str, Nothing, "text/html")
            htmlView.LinkedResources.Add(Logo1)
            htmlView.LinkedResources.Add(Logo2)

            mail = New MailMessage()
            mail.From = New MailAddress(SenderAddress)
            'mail.To.Add(Recipients)
            'mail.Bcc.Add(RecipientAddress)
            For x As Integer = 0 To recipientsList.Length - 1
                mail.Bcc.Add(recipientsList(x))
            Next
            mail.Subject = Company & " " & MailSubject & " - " & reportTime
            mail.IsBodyHtml = True
            mail.AlternateViews.Add(plainView)
            mail.AlternateViews.Add(htmlView)
            mail.Headers.Add("X-Mailer", "AlertO365 v." & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor)
            SmtpServerClient.Send(mail)
            'MsgBox("mail sent")
        Catch ex As Exception
            Console.WriteLine(Format(Now, "Short Date") & " " & Format(Now, "Long Time") & " " & ex.ToString)
        End Try

        Logo1 = Nothing
        Logo2 = Nothing
        htmlView = Nothing
        plainView = Nothing

    End Sub


    Private Shared Sub HandleError(ex As Exception)
        Console.WriteLine(ex.Message & ex.StackTrace)
    End Sub

    Private Shared Function GetEvents(Optional locale As String = "en-US") As EventInfo
        Dim requestUri As String = String.Format("{0}/{1}", serviceUrl, GetEventsMethod)
        Dim jsonDomainArray As New StringBuilder()
        Dim requestData As String = "{" & String.Format(GetEventsRequestData, cookie, "0", locale, pastDays) & "}"
        Return GetResponse(Of EventInfo)(requestUri, requestData)
    End Function

    Private Shared Function GetServiceInformation(Optional locale As String = "en-US") As ServiceInformation()
        Dim requestUri As String = String.Format("{0}/{1}", serviceUrl, GetServiceInfoMethod)
        Dim requestData As String = "{" & String.Format(GetServiceInfoRequestData, cookie, locale) & "}"
        Return GetResponse(Of ServiceInformation())(requestUri, requestData)
    End Function

    Private Shared Function GetResponse(Of T)(requestUri As String, requestData As String) As T
        ServicePointManager.ServerCertificateValidationCallback = Function(a, b, c, d)
                                                                      Return True

                                                                  End Function
        Dim data As T = Nothing

        Try
            Dim webRequest__1 As HttpWebRequest = DirectCast(WebRequest.Create(requestUri), HttpWebRequest)
            Dim postBytes As Byte() = Encoding.UTF8.GetBytes(requestData)
            webRequest__1.Method = "POST"
            webRequest__1.ContentType = "application/json"

            Dim ar As IAsyncResult = webRequest__1.BeginGetRequestStream(AddressOf ProcessWebRequest, Nothing)
            While Not ar.IsCompleted
                Thread.SpinWait(10)
            End While

            Dim startTime As DateTime = DateTime.Now

            Using requestStream As Stream = webRequest__1.EndGetRequestStream(ar)
                requestStream.Write(postBytes, 0, postBytes.Length)
            End Using

            ar = webRequest__1.BeginGetResponse(AddressOf ProcessWebResponse, Nothing)
            While Not ar.IsCompleted
                Thread.SpinWait(10)
            End While

            Dim endTime As DateTime = DateTime.Now
            Dim ts As TimeSpan = endTime - startTime
            Using response As WebResponse = webRequest__1.EndGetResponse(ar)
                Dim responseStream As Stream = response.GetResponseStream()
                Dim serailizer As New DataContractJsonSerializer(GetType(T))

                data = DirectCast(serailizer.ReadObject(responseStream), T)
            End Using
        Catch ex As Exception
            Console.WriteLine(String.Format("The following error occured :" & vbLf & "{0}", ex.Message))
        End Try

        Return data
    End Function

    Private Shared Function RegisterForSHDAccess(userName As String, password As String) As RegistrationInfo
        Dim responseData As String = String.Empty
        Dim jsonDomainArray As New StringBuilder()
        Dim requestUri As String = String.Format("{0}/{1}", serviceUrl, RegisterMethod)
        Dim requestData As String = "{" & String.Format(RegisterRequestData, userName, password) & "}"
        Dim info As RegistrationInfo = GetResponse(Of RegistrationInfo)(requestUri, requestData)
        Return info
    End Function

    Private Shared Sub ProcessWebResponse(ar As IAsyncResult)
    End Sub

    Private Shared Sub ProcessWebRequest(ar As IAsyncResult)
    End Sub

End Class


