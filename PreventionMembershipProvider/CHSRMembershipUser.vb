Imports System
Imports System.Web.Security

Public Class CHsRMembershipUser
    Inherits MembershipUser

    Private _program As Integer

    Public Property program() As Integer
        Get
            Return _program
        End Get
        Set(ByVal value As Integer)
            _program = value
        End Set
    End Property

    Private _programsite As Integer

    Public Property programsite() As Integer
        Get
            Return _programsite
        End Get
        Set(ByVal value As Integer)
            _programsite = value
        End Set
    End Property

    Private _role As String

    Public Property Role() As String
        Get
            Return _role
        End Get
        Set(ByVal value As String)
            _role = value
        End Set
    End Property

    Private _isVerified As Boolean

    Public Property IsVerified() As Boolean
        Get
            Return _isVerified
        End Get
        Set(value As Boolean)
            _isVerified = value
        End Set
    End Property

    Private _verifiedDate As DateTime

    Public Property VerifiedDate() As DateTime
        Get
            Return _verifiedDate
        End Get
        Set(value As DateTime)
            _verifiedDate = value
        End Set
    End Property

    Private _verifyCode As String

    Public Property VerifyCode() As String
        Get
            Return _verifyCode
        End Get
        Set(value As String)
            _verifyCode = value
        End Set
    End Property

    Public Sub New(ByVal providername As String,
                   ByVal username As String,
                   ByVal userid As Object,
                   ByVal email As String,
                   ByVal passwordQuestion As String,
                   ByVal comment As String,
                   ByVal isApproved As Boolean,
                   ByVal isLockedOut As Boolean,
                   ByVal creationDate As DateTime,
                   ByVal lastLoginDate As DateTime,
                   ByVal lastActivityDate As DateTime,
                   ByVal lastPasswordChangedDate As DateTime,
                   ByVal lastLockedOutDate As DateTime,
                   ByVal program As Integer,
                   ByVal role As String,
                   ByVal programsite As Integer,
                   ByVal isVerified As Boolean,
                   ByVal verifiedDate As DateTime,
                   ByVal verifyCode As String)

        MyBase.New(providername,
                   username,
                   userid,
                   email,
                   passwordQuestion,
                   comment,
                   isApproved,
                   isLockedOut,
                   creationDate,
                   lastLoginDate,
                   lastActivityDate,
                   lastPasswordChangedDate,
                   lastLockedOutDate)

        Me.program = program
        Me.Role = role
        Me.programsite = programsite

        Me.IsVerified = isVerified
        Me.VerifiedDate = verifiedDate
        Me.VerifyCode = verifyCode

    End Sub

End Class

