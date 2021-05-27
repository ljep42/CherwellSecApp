
Imports IO.Swagger.Api
Imports IO.Swagger.Model

Module Module1

    Sub Main()

        Dim ProdURL As String
        Dim key As String
        Dim user As String
        Dim pwd As String

        ProdURL = ""
        'DevURL = ""
        key = ""
        'Devkey = ""
        user = ""
        pwd = ""

        'GetTeamUserCounts(ProdURL, key, user, pwd)
        'GetUserNames(ProdURL, key, user, pwd)
        'GetUserTeams(ProdURL, key, user, pwd)
        GetSecRights(ProdURL, key, user, pwd)
        'GetSecurityGroups()

    End Sub

    Public Function GetTeamUserCounts(URL As String, key As String, user As String, pwd As String) ' gets counts of users who have team set as default in the team info table


        Dim serviceApi = New ServiceApi(URL)
        Dim tokenResponse = serviceApi.ServiceToken("password", key, Nothing, user, pwd, Nothing, "Internal")

        Dim TeamsAPI = New TeamsApi(URL)
        TeamsAPI.Configuration.AddDefaultHeader("Authorization", "Bearer " & tokenResponse.AccessToken)
        Dim Teams = TeamsAPI.TeamsGetTeamsV2()

        Dim UserAPI = New UsersApi(URL)
        UserAPI.Configuration.AddDefaultHeader("Authorization", "Bearer " & tokenResponse.AccessToken)
        Dim Users = UserAPI.UsersGetListOfUsers("Both", True)

        Dim TeamCount As New List(Of String)

        For Each t In Teams.Teams
            Dim teamName = t.TeamName
            Dim i As Integer = 0

            For Each u In Users.Users
                If u.Fields(18).Value = teamName Then
                    i += 1
                End If
            Next

            TeamCount.Add(teamName & "," & i)
        Next

        Using sw As New System.IO.StreamWriter("teamCountPROD.csv")
            For Each item As String In TeamCount
                sw.WriteLine(item)

            Next
            sw.Flush()
            sw.Close()

        End Using

        Return "done"

    End Function

    Public Function GetUserNames(URL As String, key As String, user As String, pwd As String) ' gets list of usernames both internal and windows users and outputs to file

        Dim strenvusername = System.Security.Principal.WindowsIdentity.GetCurrent().Name

        Dim serviceApi = New ServiceApi(URL)
        Dim tokenResponse = serviceApi.ServiceToken("password", key, Nothing, user, pwd, Nothing, "Internal")

        Dim UsersApi = New UsersApi(URL)
        UsersApi.Configuration.AddDefaultHeader("Authorization", "Bearer " & tokenResponse.AccessToken)

        Dim userList = UsersApi.UsersGetListOfUsers("Internal", True)
        Dim ExUserList = UsersApi.UsersGetListOfUsers("Windows", True)

        Dim uList As New List(Of String)
        uList.Add("DisplayName, isLocked, Type, SamAccount")

        For Each u In userList.Users
            Dim DisplayName = u.PublicId
            Dim locked = u.AccountLocked

            uList.Add(DisplayName & ", " & locked & ", " & "Internal")

        Next

        For Each u In ExUserList.Users
            Dim DisplayName = u.PublicId
            Dim locked = u.AccountLocked

            For Each f In u.Fields
                If f.Name = "SAMAccountName" Then
                    Dim SamAccount = f.Value
                    uList.Add(DisplayName & ", " & locked & ", " & "Windows" & ", " & SamAccount)
                End If
            Next
        Next

        Using sw As New System.IO.StreamWriter("\\shbstaffrdb\temp\UserListPROD.csv")
            For Each item As String In uList
                sw.WriteLine(item)

            Next
            sw.Flush()
            sw.Close()

        End Using

        Return "done"

    End Function

    Public Function GetUserTeams(URL As String, key As String, user As String, pwd As String) ' gets login names and team name for every user

        Dim serviceApi = New ServiceApi(URL)
        Dim tokenResponse = serviceApi.ServiceToken("password", key, Nothing, user, pwd, Nothing, "Internal")

        Dim TeamsAPI = New TeamsApi(URL)
        TeamsAPI.Configuration.AddDefaultHeader("Authorization", "Bearer " & tokenResponse.AccessToken)
        Dim Teams = TeamsAPI.TeamsGetTeamsV2()

        Dim UserAPI = New UsersApi(URL)
        UserAPI.Configuration.AddDefaultHeader("Authorization", "Bearer " & tokenResponse.AccessToken)

        Dim userList = UserAPI.UsersGetListOfUsers("Both", True)

        Dim TeamList As New List(Of String)
        TeamList.Add("SamAccount, FullName, TeamName")

        For Each u In userList.Users
            Dim userID = u.RecordId
            Dim LoginDisplayName = u.DisplayName
            Dim Fullname = u.PublicId

            'Dim res As String = ""
            'Dim idx As Integer = 0
            'For Each f In u.Fields
            '    Dim value2 = f.Name
            '    Dim value3 = f.Value

            '    res &= idx & ", " & value2 & ", " & value3 & vbNewLine
            '    idx += 1

            'Next
            'MsgBox(u.Fields)

            Dim userTeam = TeamsAPI.TeamsGetUsersTeamsV2(userID)

            For Each t In userTeam.Teams
                Dim teamName = t.TeamName
                TeamList.Add(LoginDisplayName & "," & Fullname & "," & teamName)
            Next
        Next

        Using sw As New System.IO.StreamWriter("UserteamlistPROD.csv")
            For Each item As String In TeamList
                sw.WriteLine(item)

            Next
            sw.Flush()
            sw.Close()

        End Using

        Return "done"

    End Function

    Public Function GetSecRights(URL As String, key As String, user As String, pwd As String) 'gets security rights 
        Dim serviceApi = New ServiceApi(URL)
        Dim tokenResponse = serviceApi.ServiceToken("password", key, Nothing, user, pwd, Nothing, "Internal")

        Dim SecurityAPI = New SecurityApi(URL)
        SecurityAPI.Configuration.AddDefaultHeader("Authorization", "Bearer " & tokenResponse.AccessToken)

        Dim Groups = SecurityAPI.SecurityGetSecurityGroupsV2()
        Dim Categories = SecurityAPI.SecurityGetSecurityGroupCategoriesV2()

        Dim listOfRights As New List(Of String)
        listOfRights.Add("GROUP, CATEGORY, RIGHT_NAME, ALLOW, VIEWRUNOPEN, ADD, EDIT, DELETE, YESNORIGHT")

        For Each g In Groups.SecurityGroups
            Dim groupID = g.GroupId
            Dim groupName = g.GroupName

            For Each c In Categories.RightCategories
                Dim categoryID = c.CategoryId
                Dim categoryName = c.CategoryName

                Dim secRight = SecurityAPI.SecurityGetSecurityGroupRightsByGroupIdAndCategoryIdV1(groupID, categoryID)

                For Each s In secRight
                    listOfRights.Add(groupName & "," & s.CategoryName & "," & s.RightName & "," & s.Allow & "," & s.ViewRunOpen & "," & s.Add & "," & s.Edit & "," & s.Delete & "," & s.IsYesNoRight)

                Next
            Next
        Next

        Using sw As New System.IO.StreamWriter("rightslistPROD.csv")
            For Each item As String In listOfRights
                sw.WriteLine(item)

            Next
            sw.Flush()
            sw.Close()

        End Using

        Return "done"

    End Function

    Public Function GetSecurityGroups()
        Dim serviceApi = New ServiceApi("https://lbhservicedesk.cherwellondemand.com/CherwellApi/")
        Dim tokenResponse = serviceApi.ServiceToken("password", "be46788a-d4d0-42a3-b742-bb25cee7ee83", Nothing, "cherwelltest", "cherwell@12", Nothing, "Internal")

        Dim SecurityAPI = New SecurityApi("https://lbhservicedesk.cherwellondemand.com/CherwellApi/")
        SecurityAPI.Configuration.AddDefaultHeader("Authorization", "Bearer " & tokenResponse.AccessToken)

        Dim Groups = SecurityAPI.SecurityGetSecurityGroupsV2()

        Dim listofuser As New List(Of String)
        listofuser.Add("GROUPNAME, LOGIN, DISPLAYNAME, DEFAULTTEAM")

        For Each ObjGroup In Groups.SecurityGroups

            Dim GroupID = ObjGroup.GroupId
            Dim GroupName = ObjGroup.GroupName

            Dim users = SecurityAPI.SecurityGetUsersInSecurityGroupV2(GroupID)

            For Each user In users.Users

                Dim res As String = ""
                Dim idx As Integer = 0
                For Each userfield In user.Fields
                    Dim value2 = userfield.Name
                    Dim value3 = userfield.Value

                    res &= idx & ", " & value2 & ", " & value3 & vbNewLine
                    idx += 1

                Next

                MsgBox(res)

                listofuser.Add(GroupName & "," & user.DisplayName & "," & user.Fields(7).Value & "," & user.Fields(18).Value)

            Next
        Next

        Using sw As New System.IO.StreamWriter("userSecGroupsPROD.csv")
            For Each item As String In listofuser
                sw.WriteLine(item)

            Next
            sw.Flush()
            sw.Close()

        End Using
        Return "done"

    End Function

End Module
