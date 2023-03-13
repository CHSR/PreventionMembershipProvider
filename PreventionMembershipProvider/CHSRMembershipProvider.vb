Imports System.Configuration
Imports System.Collections.Specialized
Imports System.Configuration.Provider
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Text
Imports System.Web.Configuration
Imports System.Web.Security

'
'
' This provider works with the following schema for the table of user data.
' 
'CREATE TABLE [dbo].[Users](
'	[PKID] [uniqueidentifier] NOT NULL,
'	[Username] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
'	[ApplicationName] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
'	[Email] [varchar](128) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
'	[Comment] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
'	[Password] [varchar](128) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
'	[PasswordQuestion] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
'	[PasswordAnswer] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
'	[IsApproved] [bit] NULL,
'	[LastActivityDate] [datetime] NULL,
'	[LastLoginDate] [datetime] NULL,
'	[LastPasswordChangedDate] [datetime] NULL,
'	[CreationDate] [datetime] NULL,
'	[IsOnLine] [bit] NULL,
'	[IsLockedOut] [bit] NULL,
'	[LastLockedOutDate] [datetime] NULL,
'	[FailedPasswordAttemptCount] [int] NULL,
'	[FailedPasswordAttemptWindowStart] [datetime] NULL,
'	[FailedPasswordAnswerAttemptCount] [int] NULL,
'	[FailedPasswordAnswerAttemptWindowStart] [datetime] NULL,
'	[program] [int] NULL,
'	[Role] varchar(255) NULL,
'   [isVerified] bit NOT NULL DEFAULT(0),
'   [verifiedDate] datetime NULL,
'   [verifyCode] varchar(128) NULL
'PRIMARY KEY CLUSTERED 
'(
'	[PKID] ASC
')WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
') ON [PRIMARY]
' 

Public Class CHSRMembershipProvider
    Inherits MembershipProvider

    '
    ' Global generated password length, generic exception message, event log info.
    '

    Private newPasswordLength As Integer = 8
    Private eventSource As String = "CHSRMembershipProvider"
    Private eventLog As String = "Application"
    Private exceptionMessage As String = "An exception occurred. Please check the Event Log."
    Private connectionString As String

    '
    ' Used when determining encryption key values.
    '

    Private machineKey As MachineKeySection


    '
    ' If False, exceptions are thrown to the caller. If True,
    ' exceptions are written to the event log.
    '

    Private pWriteExceptionsToEventLog As Boolean

    Public Property WriteExceptionsToEventLog() As Boolean
        Get
            Return pWriteExceptionsToEventLog
        End Get
        Set(ByVal value As Boolean)
            pWriteExceptionsToEventLog = value
        End Set
    End Property



    '
    ' System.Configuration.Provider.ProviderBase.Initialize Method
    '

    Public Overrides Sub Initialize(ByVal name As String, ByVal config As NameValueCollection)


        '
        ' Initialize values from web.config.
        '

        If config Is Nothing Then _
          Throw New ArgumentNullException("config")

        If name Is Nothing OrElse name.Length = 0 Then _
          name = "OdbcMembershipProvider"

        If String.IsNullOrEmpty(config("description")) Then
            config.Remove("description")
            config.Add("description", "Sample ODBC Membership provider")
        End If

        ' Initialize the abstract base class.
        MyBase.Initialize(name, config)


        pApplicationName = GetConfigValue(config("applicationName"), _
                                        System.Web.Hosting.HostingEnvironment.ApplicationVirtualPath)
        pMaxInvalidPasswordAttempts = Convert.ToInt32(GetConfigValue(config("maxInvalidPasswordAttempts"), "5"))
        pPasswordAttemptWindow = Convert.ToInt32(GetConfigValue(config("passwordAttemptWindow"), "10"))
        pMinRequiredNonAlphanumericCharacters = Convert.ToInt32(GetConfigValue(config("minRequiredAlphaNumericCharacters"), "1"))
        pMinRequiredPasswordLength = Convert.ToInt32(GetConfigValue(config("minRequiredPasswordLength"), "7"))
        pPasswordStrengthRegularExpression = Convert.ToString(GetConfigValue(config("passwordStrengthRegularExpression"), ""))
        pEnablePasswordReset = Convert.ToBoolean(GetConfigValue(config("enablePasswordReset"), "True"))
        pEnablePasswordRetrieval = Convert.ToBoolean(GetConfigValue(config("enablePasswordRetrieval"), "True"))
        pRequiresQuestionAndAnswer = Convert.ToBoolean(GetConfigValue(config("requiresQuestionAndAnswer"), "False"))
        pRequiresUniqueEmail = Convert.ToBoolean(GetConfigValue(config("requiresUniqueEmail"), "True"))
        pWriteExceptionsToEventLog = Convert.ToBoolean(GetConfigValue(config("writeExceptionsToEventLog"), "True"))

        Dim temp_format As String = config("passwordFormat")
        If temp_format Is Nothing Then
            temp_format = "Hashed"
        End If

        Select Case temp_format
            Case "Hashed"
                pPasswordFormat = MembershipPasswordFormat.Hashed
            Case "Encrypted"
                pPasswordFormat = MembershipPasswordFormat.Encrypted
            Case "Clear"
                pPasswordFormat = MembershipPasswordFormat.Clear
            Case Else
                Throw New ProviderException("Password format not supported.")
        End Select

        '
        ' Initialize SqlConnection.
        '

        Dim ConnectionStringSettings As ConnectionStringSettings = _
          ConfigurationManager.ConnectionStrings(config("connectionStringName"))

        If ConnectionStringSettings Is Nothing OrElse ConnectionStringSettings.ConnectionString.Trim() = "" Then
            Throw New ProviderException("Connection string cannot be blank.")
        End If

        connectionString = ConnectionStringSettings.ConnectionString


        ' Get encryption and decryption key information from the configuration.
        Dim cfg As System.Configuration.Configuration = _
          WebConfigurationManager.OpenWebConfiguration(System.Web.Hosting.HostingEnvironment.ApplicationVirtualPath)
        machineKey = CType(cfg.GetSection("system.web/machineKey"), MachineKeySection)

        If machineKey.ValidationKey.Contains("AutoGenerate") Then _
          If PasswordFormat <> MembershipPasswordFormat.Clear Then _
            Throw New ProviderException("Hashed or Encrypted passwords " & _
                                        "are not supported with auto-generated keys.")
    End Sub


    '
    ' A helper function to retrieve config values from the configuration file.
    '

    Private Function GetConfigValue(ByVal configValue As String, ByVal defaultValue As String) As String
        If String.IsNullOrEmpty(configValue) Then _
          Return defaultValue

        Return configValue
    End Function


    '
    ' System.Web.Security.MembershipProvider properties.
    '


    Private pApplicationName As String
    Private pEnablePasswordReset As Boolean
    Private pEnablePasswordRetrieval As Boolean
    Private pRequiresQuestionAndAnswer As Boolean
    Private pRequiresUniqueEmail As Boolean
    Private pMaxInvalidPasswordAttempts As Integer
    Private pPasswordAttemptWindow As Integer
    Private pPasswordFormat As MembershipPasswordFormat

    Public Overrides Property ApplicationName() As String
        Get
            Return pApplicationName
        End Get
        Set(ByVal value As String)
            pApplicationName = value
        End Set
    End Property

    Public Overrides ReadOnly Property EnablePasswordReset() As Boolean
        Get
            Return pEnablePasswordReset
        End Get
    End Property


    Public Overrides ReadOnly Property EnablePasswordRetrieval() As Boolean
        Get
            Return pEnablePasswordRetrieval
        End Get
    End Property


    Public Overrides ReadOnly Property RequiresQuestionAndAnswer() As Boolean
        Get
            Return pRequiresQuestionAndAnswer
        End Get
    End Property


    Public Overrides ReadOnly Property RequiresUniqueEmail() As Boolean
        Get
            Return pRequiresUniqueEmail
        End Get
    End Property


    Public Overrides ReadOnly Property MaxInvalidPasswordAttempts() As Integer
        Get
            Return pMaxInvalidPasswordAttempts
        End Get
    End Property


    Public Overrides ReadOnly Property PasswordAttemptWindow() As Integer
        Get
            Return pPasswordAttemptWindow
        End Get
    End Property


    Public Overrides ReadOnly Property PasswordFormat() As MembershipPasswordFormat
        Get
            Return pPasswordFormat
        End Get
    End Property

    Private pMinRequiredNonAlphanumericCharacters As Integer

    Public Overrides ReadOnly Property MinRequiredNonAlphanumericCharacters() As Integer
        Get
            Return pMinRequiredNonAlphanumericCharacters
        End Get
    End Property

    Private pMinRequiredPasswordLength As Integer

    Public Overrides ReadOnly Property MinRequiredPasswordLength() As Integer
        Get
            Return pMinRequiredPasswordLength
        End Get
    End Property

    Private pPasswordStrengthRegularExpression As String

    Public Overrides ReadOnly Property PasswordStrengthRegularExpression() As String
        Get
            Return pPasswordStrengthRegularExpression
        End Get
    End Property

    '
    ' System.Web.Security.MembershipProvider methods.
    '

    '
    ' MembershipProvider.ChangePassword
    '

    Public Overrides Function ChangePassword(ByVal username As String, _
                                             ByVal oldPwd As String, _
                                             ByVal newPwd As String) As Boolean
        If Not ValidateUser(username, oldPwd) Then _
          Return False


        Dim args As ValidatePasswordEventArgs = _
          New ValidatePasswordEventArgs(username, newPwd, True)

        OnValidatingPassword(args)

        If args.Cancel Then
            If Not args.FailureInformation Is Nothing Then
                Throw args.FailureInformation
            Else
                Throw New ProviderException("Change password canceled due to New password validation failure.")
            End If
        End If


        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("UPDATE Users " & _
          " SET Password = @Password, LastPasswordChangedDate = @LastPasswordChangedDate " & _
          " WHERE Username = @Username AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@Password", SqlDbType.VarChar, 255).Value = EncodePassword(newPwd)
        cmd.Parameters.Add("@LastPasswordChangedDate", SqlDbType.DateTime).Value = DateTime.Now
        cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName


        Dim rowsAffected As Integer = 0

        Try
            conn.Open()

            rowsAffected = cmd.ExecuteNonQuery()
        Catch e As SqlException
            If WriteExceptionsToEventLog Then

                WriteToEventLog(e, "ChangePassword")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

        If rowsAffected > 0 Then
            Return True
        End If

        Return False
    End Function



    '
    ' MembershipProvider.ChangePasswordQuestionAndAnswer
    '

    Public Overrides Function ChangePasswordQuestionAndAnswer(ByVal username As String, _
                  ByVal password As String, _
                  ByVal newPwdQuestion As String, _
                  ByVal newPwdAnswer As String) As Boolean

        If Not ValidateUser(username, password) Then _
          Return False

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("UPDATE Users " & _
                " SET PasswordQuestion = @Question, PasswordAnswer = @Answer" & _
                " WHERE Username = @Username AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@Question", SqlDbType.VarChar, 255).Value = newPwdQuestion
        cmd.Parameters.Add("@Answer", SqlDbType.VarChar, 255).Value = EncodePassword(newPwdAnswer)
        cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName


        Dim rowsAffected As Integer = 0

        Try
            conn.Open()

            rowsAffected = cmd.ExecuteNonQuery()
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "ChangePasswordQuestionAndAnswer")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

        If rowsAffected > 0 Then
            Return True
        End If

        Return False
    End Function



    '
    ' MembershipProvider.CreateUser
    '

    Public Overrides Function CreateUser(ByVal username As String, _
                                         ByVal password As String, _
                                         ByVal email As String, _
                                         ByVal passwordQuestion As String, _
                                         ByVal passwordAnswer As String, _
                                         ByVal isApproved As Boolean, _
                                         ByVal providerUserKey As Object, _
                                         ByRef status As MembershipCreateStatus) _
                              As MembershipUser
        Return Me.CreateUser(username, password, email, _
                             If(Not String.IsNullOrEmpty(passwordQuestion), passwordQuestion, String.Empty), _
                             If(Not String.IsNullOrEmpty(passwordAnswer), passwordAnswer, String.Empty), _
                             isApproved, providerUserKey, Nothing, String.Empty, status)
    End Function


    '
    ' CHSRMembershipProvider.CreateUser -- returns CHSRMembershipUser
    '

    Public Overloads Function CreateUser(ByVal username As String, _
                                         ByVal password As String, _
                                         ByVal email As String, _
                                         ByVal passwordQuestion As String, _
                                         ByVal passwordAnswer As String, _
                                         ByVal isApproved As Boolean, _
                                         ByVal providerUserKey As Object, _
                                         ByVal program As Integer, _
                                         ByVal role As String, _
                                         ByRef status As MembershipCreateStatus) _
                              As CHsRMembershipUser

        Dim Args As ValidatePasswordEventArgs = _
          New ValidatePasswordEventArgs(username, password, True)

        OnValidatingPassword(Args)

        If Args.Cancel Then
            status = MembershipCreateStatus.InvalidPassword
            Return Nothing
        End If


        If RequiresUniqueEmail AndAlso GetUserNameByEmail(email) <> "" Then
            status = MembershipCreateStatus.DuplicateEmail
            Return Nothing
        End If

        Dim u As MembershipUser = GetUser(username, False)

        If u Is Nothing Then
            Dim createDate As DateTime = DateTime.Now

            If providerUserKey Is Nothing Then
                providerUserKey = Guid.NewGuid()
            Else
                If Not TypeOf providerUserKey Is Guid Then
                    status = MembershipCreateStatus.InvalidProviderUserKey
                    Return Nothing
                End If
            End If

            Dim conn As SqlConnection = New SqlConnection(connectionString)
            Dim cmd As SqlCommand = New SqlCommand("INSERT INTO Users " & _
                   " (PKID, Username, Password, Email, PasswordQuestion, " & _
                   " PasswordAnswer, IsApproved," & _
                   " Comment, CreationDate, LastPasswordChangedDate, LastActivityDate," & _
                   " ApplicationName, IsLockedOut, LastLockedOutDate," & _
                   " FailedPasswordAttemptCount, FailedPasswordAttemptWindowStart, " & _
                   " FailedPasswordAnswerAttemptCount, FailedPasswordAnswerAttemptWindowStart, " & _
                   " program, role, isVerified)" & _
                   " Values(@PKID, @Username, @Password, @Email, @PasswordQuestion, @PasswordAnswer, @IsApproved, @Comment, @CreationDate, @LastPasswordChangedDate, @LastActivityDate, @ApplicationName, @IsLockedOut, @LastLockedOutDate, @FailedPasswordAttemptCount, @FailedPasswordAttemptWindowStart, @FailedPasswordAnswerAttemptCount, @FailedPasswordAnswerAttemptWindowStart, @program, @role, @isVerified)", conn)

            cmd.Parameters.Add("@PKID", SqlDbType.UniqueIdentifier).Value = providerUserKey
            cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
            cmd.Parameters.Add("@Password", SqlDbType.VarChar, 255).Value = EncodePassword(password)
            cmd.Parameters.Add("@Email", SqlDbType.VarChar, 128).Value = email
            cmd.Parameters.Add("@PasswordQuestion", SqlDbType.VarChar, 255).Value = passwordQuestion
            cmd.Parameters.Add("@PasswordAnswer", SqlDbType.VarChar, 255).Value = EncodePassword(passwordAnswer)
            cmd.Parameters.Add("@IsApproved", SqlDbType.Bit).Value = isApproved
            cmd.Parameters.Add("@Comment", SqlDbType.VarChar, 255).Value = ""
            cmd.Parameters.Add("@CreationDate", SqlDbType.DateTime).Value = createDate
            cmd.Parameters.Add("@LastPasswordChangedDate", SqlDbType.DateTime).Value = createDate
            cmd.Parameters.Add("@LastActivityDate", SqlDbType.DateTime).Value = createDate
            cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName
            cmd.Parameters.Add("@IsLockedOut", SqlDbType.Bit).Value = False
            cmd.Parameters.Add("@LastLockedOutDate", SqlDbType.DateTime).Value = createDate
            cmd.Parameters.Add("@FailedPasswordAttemptCount", SqlDbType.Int).Value = 0
            cmd.Parameters.Add("@FailedPasswordAttemptWindowStart", SqlDbType.DateTime).Value = createDate
            cmd.Parameters.Add("@FailedPasswordAnswerAttemptCount", SqlDbType.Int).Value = 0
            cmd.Parameters.Add("@FailedPasswordAnswerAttemptWindowStart", SqlDbType.DateTime).Value = createDate
            cmd.Parameters.Add("@program", SqlDbType.Int).Value = program
            cmd.Parameters.Add("@role", SqlDbType.VarChar, 255).Value = role
            cmd.Parameters.Add("@isVerified", SqlDbType.Bit).Value = False

            Try
                conn.Open()

                Dim recAdded As Integer = cmd.ExecuteNonQuery()

                If recAdded > 0 Then
                    status = MembershipCreateStatus.Success
                Else
                    status = MembershipCreateStatus.UserRejected
                End If
            Catch e As SqlException
                If WriteExceptionsToEventLog Then
                    WriteToEventLog(e, "CreateUser")
                End If

                status = MembershipCreateStatus.ProviderError
            Finally
                conn.Close()
            End Try


            Return GetUser(username, False)
        Else
            status = MembershipCreateStatus.DuplicateUserName
        End If

        Return Nothing
    End Function



    '
    ' MembershipProvider.DeleteUser
    '

    Public Overrides Function DeleteUser(ByVal username As String, _
                                         ByVal deleteAllRelatedData As Boolean) As Boolean

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("DELETE FROM Users " & _
                " WHERE Username = @Username AND Applicationname = @ApplicationName", conn)

        cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

        Dim rowsAffected As Integer = 0

        Try
            conn.Open()

            rowsAffected = cmd.ExecuteNonQuery()

            If deleteAllRelatedData Then
                ' Process commands to delete all data for the user in the database.
            End If
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "DeleteUser")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

        If rowsAffected > 0 Then _
          Return True

        Return False
    End Function



    '
    ' MembershipProvider.GetAllUsers
    '

    Public Overrides Function GetAllUsers(ByVal pageIndex As Integer, _
    ByVal pageSize As Integer, _
                                          ByRef totalRecords As Integer) _
                              As MembershipUserCollection

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT Count(*) FROM Users  " & _
                                          "WHERE ApplicationName = @ApplicationName", conn)
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = ApplicationName

        Dim users As MembershipUserCollection = New MembershipUserCollection()

        Dim reader As SqlDataReader = Nothing
        totalRecords = 0

        Try
            conn.Open()
            totalRecords = CInt(cmd.ExecuteScalar())

            If totalRecords <= 0 Then Return users

            cmd.CommandText = "SELECT PKID, Username, Email, PasswordQuestion," &
                     " Comment, IsApproved, IsLockedOut, CreationDate, LastLoginDate," &
                     " LastActivityDate, LastPasswordChangedDate, LastLockedOutDate, program, role, programsite, isVerified, verifiedDate, verifyCode " &
                     " FROM Users  " &
                     " WHERE ApplicationName = @ApplicationName " &
                     " ORDER BY Username Asc"

            reader = cmd.ExecuteReader()

            Dim counter As Integer = 0
            Dim startIndex As Integer = pageSize * pageIndex
            Dim endIndex As Integer = startIndex + pageSize - 1

            Do While reader.Read()
                If counter >= startIndex Then
                    Dim u As MembershipUser = GetUserFromReader(reader)
                    users.Add(u)
                End If

                If counter >= endIndex Then cmd.Cancel()

                counter += 1
            Loop
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetAllUsers")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try

        Return users
    End Function


    '
    ' MembershipProvider.GetNumberOfUsersOnline
    '

    Public Overrides Function GetNumberOfUsersOnline() As Integer

        Dim onlineSpan As TimeSpan = New TimeSpan(0, System.Web.Security.Membership.UserIsOnlineTimeWindow, 0)
        Dim compareTime As DateTime = DateTime.Now.Subtract(onlineSpan)

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT Count(*) FROM Users " &
                " WHERE LastActivityDate > @CompareDate AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@CompareDate", SqlDbType.DateTime).Value = compareTime
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

        Dim numOnline As Integer = 0

        Try
            conn.Open()

            numOnline = CInt(cmd.ExecuteScalar())
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetNumberOfUsersOnline")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

        Return numOnline
    End Function

    '
    ' MembershipProvider.GetPassword
    '

    Public Overrides Function GetPassword(ByVal username As String, ByVal answer As String) As String

        If Not EnablePasswordRetrieval Then
            Throw New ProviderException("Password Retrieval Not Enabled.")
        End If

        If PasswordFormat = MembershipPasswordFormat.Hashed Then
            Throw New ProviderException("Cannot retrieve Hashed passwords.")
        End If

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT Password, PasswordAnswer, IsLockedOut FROM Users " & _
              " WHERE Username = @Username AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

        Dim password As String = ""
        Dim passwordAnswer As String = ""
        Dim reader As SqlDataReader = Nothing

        Try
            conn.Open()

            reader = cmd.ExecuteReader(CommandBehavior.SingleRow)

            If reader.HasRows Then
                reader.Read()

                If reader.GetBoolean(2) Then _
                  Throw New MembershipPasswordException("The supplied user is locked out.")

                password = reader.GetString(0)
                passwordAnswer = reader.GetString(1)
            Else
                Throw New MembershipPasswordException("The supplied user name is not found.")
            End If
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetPassword")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try


        If RequiresQuestionAndAnswer AndAlso Not CheckPassword(answer, passwordAnswer) Then
            UpdateFailureCount(username, "passwordAnswer")

            Throw New MembershipPasswordException("Incorrect password answer.")
        End If


        If PasswordFormat = MembershipPasswordFormat.Encrypted Then
            password = UnEncodePassword(password)
        End If

        Return password
    End Function



    '
    ' MembershipProvider.GetUser(String, Boolean)
    '

    Public Overrides Function GetUser(ByVal username As String, _
                                      ByVal userIsOnline As Boolean) As MembershipUser

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT PKID, Username, Email, PasswordQuestion," &
              " Comment, IsApproved, IsLockedOut, CreationDate, LastLoginDate," &
              " LastActivityDate, LastPasswordChangedDate, LastLockedOutDate, program, role, programsite, isVerified, verifiedDate, verifyCode" &
              " FROM Users  WHERE Username = @Username AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

        Dim u As CHsRMembershipUser = Nothing
        Dim reader As SqlDataReader = Nothing

        Try
            conn.Open()

            reader = cmd.ExecuteReader()

            If reader.HasRows Then
                reader.Read()
                u = GetUserFromReader(reader)

                If userIsOnline Then
                    Dim updateCmd As SqlCommand = New SqlCommand("UPDATE Users  " & _
                              "SET LastActivityDate = @LastActivityDate " & _
                              "WHERE Username = @Username AND Applicationname = @Applicationname", conn)

                    updateCmd.Parameters.Add("@LastActivityDate", SqlDbType.DateTime).Value = DateTime.Now
                    updateCmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
                    updateCmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

                    updateCmd.ExecuteNonQuery()
                End If
            End If
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetUser(String, Boolean)")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()

            conn.Close()
        End Try

        Return u
    End Function


    '
    ' MembershipProvider.GetUser(Object, Boolean)
    '

    Public Overrides Function GetUser(ByVal providerUserKey As Object, _
    ByVal userIsOnline As Boolean) As MembershipUser

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT PKID, Username, Email, PasswordQuestion," &
              " Comment, IsApproved, IsLockedOut, CreationDate, LastLoginDate," &
              " LastActivityDate, LastPasswordChangedDate, LastLockedOutDate, program, role, programsite, isVerified, verifiedDate, verifyCode " &
              " FROM Users  WHERE PKID = @PKID", conn)

        cmd.Parameters.Add("@PKID", SqlDbType.UniqueIdentifier).Value = providerUserKey

        Dim u As CHsRMembershipUser = Nothing
        Dim reader As SqlDataReader = Nothing

        Try
            conn.Open()

            reader = cmd.ExecuteReader()

            If reader.HasRows Then
                reader.Read()
                u = GetUserFromReader(reader)

                If userIsOnline Then
                    Dim updateCmd As SqlCommand = New SqlCommand("UPDATE Users  " & _
                              "SET LastActivityDate = @LastActivityDate " & _
                              "WHERE PKID = @PKID", conn)

                    updateCmd.Parameters.Add("@LastActivityDate", SqlDbType.DateTime).Value = DateTime.Now
                    updateCmd.Parameters.Add("@PKID", SqlDbType.UniqueIdentifier).Value = providerUserKey

                    updateCmd.ExecuteNonQuery()
                End If
            End If
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetUser(Object, Boolean)")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()

            conn.Close()
        End Try

        Return u
    End Function


    '
    ' GetUserFromReader
    '    A helper function that takes the current row from the SqlDataReader
    ' and hydrates a MembershiUser from the values. Called by the 
    ' MembershipUser.GetUser implementation.
    '

    Private Function GetUserFromReader(ByVal reader As SqlDataReader) As CHsRMembershipUser
        Dim providerUserKey As Object = reader.GetValue(0)
        Dim username As String = reader.GetString(1)
        Dim email As String = reader.GetString(2)

        Dim passwordQuestion As String = ""
        If Not reader.GetValue(3) Is DBNull.Value Then _
          passwordQuestion = reader.GetString(3)

        Dim comment As String = ""
        If Not reader.GetValue(4) Is DBNull.Value Then _
          comment = reader.GetString(4)

        Dim isApproved As Boolean = reader.GetBoolean(5)
        Dim isLockedOut As Boolean = reader.GetBoolean(6)
        Dim creationDate As DateTime = reader.GetDateTime(7)

        Dim lastLoginDate As DateTime = New DateTime()
        If Not reader.GetValue(8) Is DBNull.Value Then _
          lastLoginDate = reader.GetDateTime(8)

        Dim lastActivityDate As DateTime
        If Not reader.GetValue(9) Is DBNull.Value Then lastActivityDate = reader.GetDateTime(9)
        Dim lastPasswordChangedDate As DateTime
        If Not reader.GetValue(10) Is DBNull.Value Then lastPasswordChangedDate = reader.GetValue(10)

        Dim lastLockedOutDate As DateTime = New DateTime()
        If Not reader.GetValue(11) Is DBNull.Value Then _
          lastLockedOutDate = reader.GetDateTime(11)

        Dim program As Integer = reader.GetInt32(12)

        Dim role As String = reader.GetString(13)

        Dim programsite As Integer = reader.GetInt32(14)

        Dim isVerified As Boolean = reader.GetBoolean(15)

        Dim verifiedDate As New DateTime()
        If Not reader.GetValue(16) Is DBNull.Value Then _
          verifiedDate = reader.GetDateTime(16)

        Dim verifyCode As String
        If Not reader.GetValue(17) Is DBNull.Value Then _
            verifyCode = reader.GetString(17)

        Dim u As CHsRMembershipUser = New CHsRMembershipUser(Me.Name,
                                              username,
                                              providerUserKey,
                                              email,
                                              passwordQuestion,
                                              comment,
                                              isApproved,
                                              isLockedOut,
                                              creationDate,
                                              lastLoginDate,
                                              lastActivityDate,
                                              lastPasswordChangedDate,
                                              lastLockedOutDate,
                                              program,
                                              role,
                                              programsite,
                                              isVerified,
                                              verifiedDate,
                                              verifyCode)

        Return u
    End Function


    '
    ' MembershipProvider.UnlockUser
    '

    Public Overrides Function UnlockUser(ByVal username As String) As Boolean
        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("UPDATE Users  " & _
                                          " SET IsLockedOut = 0, LastLockedOutDate = @LastLockedOutDate " & _
                                          " WHERE Username = @Username AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@LastLockedOutDate", SqlDbType.DateTime).Value = DateTime.Now
        cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

        Dim rowsAffected As Integer = 0

        Try
            conn.Open()

            rowsAffected = cmd.ExecuteNonQuery()
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "UnlockUser")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

        If rowsAffected > 0 Then _
          Return True

        Return False
    End Function


    '
    ' MembershipProvider.GetUserNameByEmail
    '

    Public Overrides Function GetUserNameByEmail(ByVal email As String) As String
        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT Username" & _
              " FROM Users  WHERE Email = @Email AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@Email", SqlDbType.VarChar, 128).Value = email
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

        Dim username As String = ""

        Try
            conn.Open()

            username = cmd.ExecuteScalar()
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetUserNameByEmail")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

        If username Is Nothing Then _
          username = ""

        Return username
    End Function




    '
    ' MembershipProvider.ResetPassword
    '

    Public Overrides Function ResetPassword(ByVal username As String, ByVal answer As String) As String

        If Not EnablePasswordReset Then
            Throw New NotSupportedException("Password Reset is not enabled.")
        End If

        If answer Is Nothing AndAlso RequiresQuestionAndAnswer Then
            UpdateFailureCount(username, "passwordAnswer")

            Throw New ProviderException("Password answer required for password Reset.")
        End If

        Dim newPassword As String = _
          System.Web.Security.Membership.GeneratePassword(newPasswordLength, MinRequiredNonAlphanumericCharacters)


        Dim Args As ValidatePasswordEventArgs = _
          New ValidatePasswordEventArgs(username, newPassword, True)

        OnValidatingPassword(Args)

        If Args.Cancel Then
            If Not Args.FailureInformation Is Nothing Then
                Throw Args.FailureInformation
            Else
                Throw New MembershipPasswordException("Reset password canceled due to password validation failure.")
            End If
        End If


        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT PasswordAnswer, IsLockedOut FROM Users " & _
              " WHERE Username = @Username AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

        Dim rowsAffected As Integer = 0
        Dim passwordAnswer As String = ""
        Dim reader As SqlDataReader = Nothing

        Try
            conn.Open()

            reader = cmd.ExecuteReader(CommandBehavior.SingleRow)

            If reader.HasRows Then
                reader.Read()

                If reader.GetBoolean(1) Then Throw New MembershipPasswordException("The supplied user is locked out.")

                '  passwordAnswer = reader.GetString(0)
            Else
                Throw New MembershipPasswordException("The supplied user name is not found.")
            End If

            If RequiresQuestionAndAnswer AndAlso Not CheckPassword(answer, passwordAnswer) Then
                UpdateFailureCount(username, "passwordAnswer")

                Throw New MembershipPasswordException("Incorrect password answer.")
            End If

            Dim updateCmd As SqlCommand = New SqlCommand("UPDATE Users " & _
                " SET Password = @Password, LastPasswordChangedDate = @LastPasswordChangedDate" & _
                " WHERE Username = @Username AND ApplicationName = @ApplicationName AND IsLockedOut = 0", conn)

            updateCmd.Parameters.Add("@Password", SqlDbType.VarChar, 255).Value = EncodePassword(newPassword)
            updateCmd.Parameters.Add("@LastPasswordChangedDate", SqlDbType.DateTime).Value = DateTime.Now
            updateCmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
            updateCmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

            rowsAffected = updateCmd.ExecuteNonQuery()
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "ResetPassword")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try

        If rowsAffected > 0 Then
            Return newPassword
        Else
            Throw New MembershipPasswordException("User not found, or user is locked out. Password not Reset.")
        End If
    End Function


    Public Overrides Sub UpdateUser(ByVal user As MembershipUser)

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("UPDATE Users " &
                " SET Email = @Email, Comment = @Comment," &
                " IsApproved = @IsApproved, programsite = @program, role = @role, isVerified = @isVerified, verifiedDate = @verifiedDate, verifyCode= @verifyCode " &
                " WHERE Username = @Username AND ApplicationName = @ApplicationName", conn)

        Dim u As CHsRMembershipUser = CType(user, CHsRMembershipUser)

        cmd.Parameters.Add("@Email", SqlDbType.VarChar, 128).Value = user.Email
        cmd.Parameters.Add("@Comment", SqlDbType.VarChar, 255).Value = user.Comment
        cmd.Parameters.Add("@IsApproved", SqlDbType.Bit).Value = user.IsApproved
        cmd.Parameters.Add("@program", SqlDbType.Int).Value = u.program
        cmd.Parameters.Add("@role", SqlDbType.VarChar, 255).Value = u.Role
        cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = user.UserName
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName
        cmd.Parameters.Add("@isVerified", SqlDbType.Bit).Value = u.IsVerified
        cmd.Parameters.Add("@verifiedDate", SqlDbType.DateTime).Value = u.VerifiedDate
        cmd.Parameters.Add("@verifyCode", SqlDbType.VarChar, 128).Value = u.VerifyCode


        Try
            conn.Open()

            cmd.ExecuteNonQuery()
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "UpdateUser")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try
    End Sub


    '
    ' MembershipProvider.ValidateUser
    '

    Public Overrides Function ValidateUser(ByVal username As String, ByVal password As String) As Boolean
        Dim isValid As Boolean = False

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT Password, IsApproved FROM Users " &
                " WHERE Username = @Username AND ApplicationName = @ApplicationName AND IsLockedOut = 0", conn)

        cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

        Dim reader As SqlDataReader = Nothing
        Dim isApproved As Boolean = False
        Dim pwd As String = ""

        Try
            conn.Open()

            reader = cmd.ExecuteReader(CommandBehavior.SingleRow)

            If reader.HasRows Then
                reader.Read()
                pwd = reader.GetString(0)
                isApproved = reader.GetBoolean(1)
            Else
                Return False
            End If

            reader.Close()

            If CheckPassword(password, pwd) Then
                If isApproved Then
                    isValid = True

                    Dim updateCmd As SqlCommand = New SqlCommand("UPDATE Users  SET LastLoginDate = @LastLoginDate" & _
                                                            " WHERE Username = @Username AND ApplicationName = @ApplicationName", conn)

                    updateCmd.Parameters.Add("@LastLoginDate", SqlDbType.DateTime).Value = DateTime.Now
                    updateCmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
                    updateCmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

                    updateCmd.ExecuteNonQuery()
                End If
            Else
                conn.Close()

                UpdateFailureCount(username, "password")
            End If
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "ValidateUser")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try

        Return isValid
    End Function


    '
    ' UpdateFailureCount
    '   A helper method that performs the checks and updates associated with
    ' password failure tracking.
    '

    Private Sub UpdateFailureCount(ByVal username As String, ByVal failureType As String)

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT FailedPasswordAttemptCount, " & _
                                          "  FailedPasswordAttemptWindowStart, " & _
                                          "  FailedPasswordAnswerAttemptCount, " & _
                                          "  FailedPasswordAnswerAttemptWindowStart " & _
                                          "  FROM Users  " & _
                                          "  WHERE Username = @Username AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

        Dim reader As SqlDataReader = Nothing
        Dim windowStart As DateTime = New DateTime()
        Dim failureCount As Integer = 0

        Try
            conn.Open()

            reader = cmd.ExecuteReader(CommandBehavior.SingleRow)

            If reader.HasRows Then
                reader.Read()

                If failureType = "password" Then
                    If reader.IsDBNull(0) Then
                        failureCount = 0
                    Else
                        reader.GetInt32(0)
                    End If
                    If reader.IsDBNull(1) Then
                        windowStart = Date.Now()
                    Else
                        windowStart = reader.GetDateTime(1)
                    End If
                End If

                If failureType = "passwordAnswer" Then
                    failureCount = reader.GetInt32(2)
                    windowStart = reader.GetDateTime(3)
                End If
            End If

            reader.Close()

            Dim windowEnd As DateTime = windowStart.AddMinutes(PasswordAttemptWindow)

            If failureCount = 0 OrElse DateTime.Now > windowEnd Then
                ' First password failure or outside of PasswordAttemptWindow. 
                ' Start a New password failure count from 1 and a New window starting now.

                If failureType = "password" Then _
                  cmd.CommandText = "UPDATE Users  " & _
                                    "  SET FailedPasswordAttemptCount = @Count, " & _
                                    "      FailedPasswordAttemptWindowStart = @WindowStart " & _
                                    "  WHERE Username = @Username AND ApplicationName = @ApplicationName"

                If failureType = "passwordAnswer" Then _
                  cmd.CommandText = "UPDATE Users  " & _
                                    "  SET FailedPasswordAnswerAttemptCount = @Count, " & _
                                    "      FailedPasswordAnswerAttemptWindowStart = @WindowStart " & _
                                    "  WHERE Username = @Username AND ApplicationName = @ApplicationName"

                cmd.Parameters.Clear()

                cmd.Parameters.Add("@Count", SqlDbType.Int).Value = 1
                cmd.Parameters.Add("@WindowStart", SqlDbType.DateTime).Value = DateTime.Now
                cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
                cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

                If cmd.ExecuteNonQuery() < 0 Then _
                  Throw New ProviderException("Unable to update failure count and window start.")
            Else
                failureCount += 1

                If failureCount >= MaxInvalidPasswordAttempts Then
                    ' Password attempts have exceeded the failure threshold. Lock out
                    ' the user.

                    cmd.CommandText = "UPDATE Users  " & _
                                      "  SET IsLockedOut = @IsLockedOut, LastLockedOutDate = @LastLockedOutDate " & _
                                      "  WHERE Username = @Username AND ApplicationName = @ApplicationName"

                    cmd.Parameters.Clear()

                    cmd.Parameters.Add("@IsLockedOut", SqlDbType.Bit).Value = True
                    cmd.Parameters.Add("@LastLockedOutDate", SqlDbType.DateTime).Value = DateTime.Now
                    cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
                    cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

                    If cmd.ExecuteNonQuery() < 0 Then _
                      Throw New ProviderException("Unable to lock out user.")
                Else
                    ' Password attempts have not exceeded the failure threshold. Update
                    ' the failure counts. Leave the window the same.

                    If failureType = "password" Then _
                      cmd.CommandText = "UPDATE Users  " & _
                                        "  SET FailedPasswordAttemptCount = @Count" & _
                                        "  WHERE Username = @Username AND ApplicationName = @ApplicationName"

                    If failureType = "passwordAnswer" Then _
                      cmd.CommandText = "UPDATE Users  " & _
                                        "  SET FailedPasswordAnswerAttemptCount = @Count" & _
                                        "  WHERE Username = @Username AND ApplicationName = @ApplicationName"

                    cmd.Parameters.Clear()

                    cmd.Parameters.Add("@Count", SqlDbType.Int).Value = failureCount
                    cmd.Parameters.Add("@Username", SqlDbType.VarChar, 255).Value = username
                    cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

                    If cmd.ExecuteNonQuery() < 0 Then _
                      Throw New ProviderException("Unable to update failure count.")
                End If
            End If
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "UpdateFailureCount")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try
    End Sub


    '
    ' CheckPassword
    '   Compares password values based on the MembershipPasswordFormat.
    '
    '
    '   Will also be used to check the verification code in 2-factor authentication
    '
    '

    Private Function CheckPassword(ByVal password As String, ByVal dbpassword As String) As Boolean
        Dim pass1 As String = password
        Dim pass2 As String = dbpassword

        Select Case PasswordFormat
            Case MembershipPasswordFormat.Encrypted
                pass2 = UnEncodePassword(dbpassword)
            Case MembershipPasswordFormat.Hashed
                pass1 = EncodePassword(password)
            Case Else
        End Select

        If pass1 = pass2 Then
            Return True
        End If

        Return False
    End Function


    '
    ' EncodePassword
    '   Encrypts, Hashes, or leaves the password clear based on the PasswordFormat.
    '

    Private Function EncodePassword(ByVal password As String) As String

        If String.IsNullOrEmpty(password) Then
            Return password
        End If

        Dim encodedPassword As String = password

        Select Case PasswordFormat
            Case MembershipPasswordFormat.Clear

            Case MembershipPasswordFormat.Encrypted
                encodedPassword = _
                  Convert.ToBase64String(EncryptPassword(Encoding.Unicode.GetBytes(password)))
            Case MembershipPasswordFormat.Hashed
                Dim hash As HMACSHA1 = New HMACSHA1()
                hash.Key = HexToByte(machineKey.ValidationKey)
                encodedPassword = _
                  Convert.ToBase64String(hash.ComputeHash(Encoding.Unicode.GetBytes(password)))
            Case Else
                Throw New ProviderException("Unsupported password format.")
        End Select

        Return encodedPassword
    End Function


    '
    ' UnEncodePassword
    '   Decrypts or leaves the password clear based on the PasswordFormat.
    '

    Private Function UnEncodePassword(ByVal encodedPassword As String) As String
        Dim password As String = encodedPassword

        Select Case PasswordFormat
            Case MembershipPasswordFormat.Clear

            Case MembershipPasswordFormat.Encrypted
                password = _
                  Encoding.Unicode.GetString(DecryptPassword(Convert.FromBase64String(password)))
            Case MembershipPasswordFormat.Hashed
                Throw New ProviderException("Cannot unencode a hashed password.")
            Case Else
                Throw New ProviderException("Unsupported password format.")
        End Select

        Return password
    End Function

    '
    ' HexToByte
    '   Converts a hexadecimal string to a byte array. Used to convert encryption
    ' key values from the configuration.
    '

    Private Function HexToByte(ByVal hexString As String) As Byte()
        Dim ReturnBytes((hexString.Length \ 2) - 1) As Byte
        For i As Integer = 0 To ReturnBytes.Length - 1
            ReturnBytes(i) = Convert.ToByte(hexString.Substring(i * 2, 2), 16)
        Next
        Return ReturnBytes
    End Function


    '
    ' MembershipProvider.FindUsersByName
    '

    Public Overrides Function FindUsersByName(ByVal usernameToMatch As String, _
                                              ByVal pageIndex As Integer, _
                                              ByVal pageSize As Integer, _
                                              ByRef totalRecords As Integer) _
                              As MembershipUserCollection

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT Count(*) FROM Users  " & _
                  "WHERE Username LIKE @UsernameSearch AND ApplicationName = @ApplicationName", conn)
        cmd.Parameters.Add("@UsernameSearch", SqlDbType.VarChar, 255).Value = usernameToMatch
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = pApplicationName

        Dim users As MembershipUserCollection = New MembershipUserCollection()

        Dim reader As SqlDataReader = Nothing

        Try
            conn.Open()
            totalRecords = CInt(cmd.ExecuteScalar())

            If totalRecords <= 0 Then Return users

            cmd.CommandText = "SELECT PKID, Username, Email, PasswordQuestion," &
              " Comment, IsApproved, IsLockedOut, CreationDate, LastLoginDate," &
              " LastActivityDate, LastPasswordChangedDate, LastLockedOutDate, program, role, programsite, isVerified, verifiedDate, verifyCode " &
              " FROM Users  " &
              " WHERE Username LIKE @UsernameSearch AND ApplicationName = @ApplicationName " &
              " ORDER BY Username Asc"

            reader = cmd.ExecuteReader()

            Dim counter As Integer = 0
            Dim startIndex As Integer = pageSize * pageIndex
            Dim endIndex As Integer = startIndex + pageSize - 1

            Do While reader.Read()
                If counter >= startIndex Then
                    Dim u As MembershipUser = GetUserFromReader(reader)
                    users.Add(u)
                End If

                If counter >= endIndex Then cmd.Cancel()

                counter += 1
            Loop
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "FindUsersByName")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()

            conn.Close()
        End Try

        Return users
    End Function

    '
    ' MembershipProvider.FindUsersByEmail
    '

    Public Overrides Function FindUsersByEmail(ByVal emailToMatch As String, _
                                               ByVal pageIndex As Integer, _
                                               ByVal pageSize As Integer, _
                                               ByRef totalRecords As Integer) _
                              As MembershipUserCollection

        Dim conn As SqlConnection = New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand("SELECT Count(*) FROM Users  " & _
                                          "WHERE Email LIKE @EmailSearch AND ApplicationName = @ApplicationName", conn)
        cmd.Parameters.Add("@EmailSearch", SqlDbType.VarChar, 255).Value = emailToMatch
        cmd.Parameters.Add("@ApplicationName", SqlDbType.VarChar, 255).Value = ApplicationName

        Dim users As MembershipUserCollection = New MembershipUserCollection()

        Dim reader As SqlDataReader = Nothing
        totalRecords = 0

        Try
            conn.Open()
            totalRecords = CInt(cmd.ExecuteScalar())

            If totalRecords <= 0 Then Return users

            cmd.CommandText = "SELECT PKID, Username, Email, PasswordQuestion," &
                     " Comment, IsApproved, IsLockedOut, CreationDate, LastLoginDate," &
                     " LastActivityDate, LastPasswordChangedDate, LastLockedOutDate, program, role, programsite, isVerified, verifiedDate, verifyCode " &
                     " FROM Users  " &
                     " WHERE Email LIKE @EmailSearch AND ApplicationName = @ApplicationName " &
                     " ORDER BY Username Asc"

            reader = cmd.ExecuteReader()

            Dim counter As Integer = 0
            Dim startIndex As Integer = pageSize * pageIndex
            Dim endIndex As Integer = startIndex + pageSize - 1

            Do While reader.Read()
                If counter >= startIndex Then
                    Dim u As MembershipUser = GetUserFromReader(reader)
                    users.Add(u)
                End If

                If counter >= endIndex Then cmd.Cancel()

                counter += 1
            Loop
        Catch e As SqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "FindUsersByEmail")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()

            conn.Close()
        End Try

        Return users
    End Function


    '
    ' WriteToEventLog
    '   A helper function that writes exception detail to the event log. Exceptions
    ' are written to the event log as a security measure to aSub Private database
    ' details from being Returned to the browser. If a method does not Return a status
    ' or boolean indicating the action succeeded or failed, a generic exception is also 
    ' Thrown by the caller.
    '

    Private Sub WriteToEventLog(ByVal e As Exception, ByVal action As String)
        Dim log As EventLog = New EventLog()
        log.Source = eventSource
        log.Log = eventLog

        Dim message As String = "An exception occurred communicating with the data source." & vbCrLf & vbCrLf
        message &= "Action: " & action & vbCrLf & vbCrLf
        message &= "Exception: " & e.ToString()

        log.WriteEntry(message)
    End Sub


End Class