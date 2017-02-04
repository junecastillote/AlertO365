Imports System.Collections.Generic
Imports System.Text
Imports System.Threading.Tasks
Imports System.Configuration

Public Class AppConfig
	Inherits ConfigurationSection
	<ConfigurationProperty("ServiceURL")> _
	Public Property ServiceURL() As String
		Get
			Return TryCast(Me("ServiceURL"), String)
		End Get
		Set
			Me("ServiceURL") = value
		End Set
    End Property

    <ConfigurationProperty("SendViaEmail")> _
    Public Property SendViaEmail() As String
        Get
            Return TryCast(Me("SendViaEmail"), String)
        End Get
        Set(value As String)
            Me("SendViaEmail") = value
        End Set
    End Property

	<ConfigurationProperty("DomainNames")> _
	Public Property DomainNames() As String
		Get
			Return TryCast(Me("DomainNames"), String)
		End Get
		Set
			Me("DomainNames") = value
		End Set
	End Property

	<ConfigurationProperty("UserName")> _
	Public Property UserName() As String
		Get
			Return TryCast(Me("UserName"), String)
		End Get
		Set
			Me("UserName") = value
		End Set
	End Property

	<ConfigurationProperty("Password")> _
	Public Property Password() As String
		Get
			Return TryCast(Me("Password"), String)
		End Get
		Set
			Me("Password") = value
		End Set
	End Property

	<ConfigurationProperty("IsAOBO")> _
	Public Property IsAOBO() As String
		Get
			Return TryCast(Me("IsAOBO"), String)
		End Get
		Set
			Me("IsAOBO") = value
		End Set
	End Property

	<ConfigurationProperty("PastDays")> _
	Public Property PastDays() As String
		Get
			Return TryCast(Me("PastDays"), String)
		End Get
		Set
			Me("PastDays") = value
		End Set
    End Property

    <ConfigurationProperty("SMTPServer")> _
    Public Property SMTPServer() As String
        Get
            Return TryCast(Me("SMTPServer"), String)
        End Get
        Set(value As String)
            Me("SMTPServer") = value
        End Set
    End Property

    <ConfigurationProperty("SenderAddress")> _
    Public Property SenderAddress() As String
        Get
            Return TryCast(Me("SenderAddress"), String)
        End Get
        Set(value As String)
            Me("SenderAddress") = value
        End Set
    End Property

    <ConfigurationProperty("RecipientAddress")> _
    Public Property RecipientAddress() As String
        Get
            Return TryCast(Me("RecipientAddress"), String)
        End Get
        Set(value As String)
            Me("RecipientAddress") = value
        End Set
    End Property

    <ConfigurationProperty("MailSubject")> _
    Public Property MailSubject() As String
        Get
            Return TryCast(Me("MailSubject"), String)
        End Get
        Set(value As String)
            Me("MailSubject") = value
        End Set
    End Property

    <ConfigurationProperty("Company")> _
    Public Property Company() As String
        Get
            Return TryCast(Me("Company"), String)
        End Get
        Set(value As String)
            Me("Company") = value
        End Set
    End Property
End Class
