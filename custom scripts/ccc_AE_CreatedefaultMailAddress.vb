' ccc_AE_CreatedefaultMailAddress
' Returns the default email address of a person.
#If Not SCRIPTDEBUGGER Then
Imports System.Collections.Generic
Imports System.Data
#End If

Public Overridable Function ccc_VI_AE_CreatedefaultMailAddress(ByVal Lastname As String, ByVal Firstname As String, ByVal uidperson As String) As String
    Dim i As Int32 = 1
    
    Firstname = ccc_Translit(Firstname)
    Dim FirstnameLen As Int32 = Len(Firstname)
    Lastname = ccc_Translit(Lastname)
    
    Dim template As String = "." & Lastname
    
    Dim prefix As String
    Dim index As String
    Dim defaultemailaddress As String
    
    Dim maildom As String = String.Empty
    Dim f As ISqlFormatter = Connection.SqlFormatter

    ccc_VI_AE_CreatedefaultMailAddress = String.Empty

    maildom = Connection.GetConfigParm("QER\Person\DefaultMailDomain")

    If maildom = "" Then
        Exit Function
    Else
        If Not maildom.StartsWith("@") Then
            maildom = "@" & maildom
        End If
    End If

    ' fill existing addresses in a dictionary
    Dim existing As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
    Dim dummy As New Object()
    Dim dummyPerson As ISingleDbObject
    dummyPerson = Connection.CreateSingle("Person")
    Dim pattern As String = "%" & template & "%" & maildom
    Dim myObjectKey As New DbObjectKey("Person", uidperson)

    Using rd As IDataReader = CType(dummyPerson.Custom.CallMethod("SearchMailAddresses", pattern), IDataReader)

        While rd.Read()
            Dim address As String
            Dim objectKeyString As String
            Dim objectKey As DbObjectKey

            address = rd.GetString(rd.GetOrdinal("smtp"))
            objectKeyString = rd.GetString(rd.GetOrdinal("ObjectKeyPerson"))

            If Not String.IsNullOrEmpty(objectKeyString) Then
                objectKey = New DbObjectKey(objectKeyString)

                ' only addresses which not belong to the actual employee will be considered
                If myObjectKey.Equals(objectKey) Then
                    Continue While
                End If
            End If

            existing(address) = dummy
        End While
    End Using

    While True

        If i <= FirstnameLen
            prefix = Firstname.Substring(0, i)
            defaultemailaddress = prefix & template & maildom
        Else
            index = CStr(i - FirstnameLen)
            defaultemailaddress = Firstname & template & index & maildom
        End If
        
        ' Does not exists?
        If Not existing.ContainsKey(defaultemailaddress) Then
            Return defaultemailaddress
        End If
        
        ' next trial
        i = i + 1

    End While
End Function