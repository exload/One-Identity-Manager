' ccc_AE_BuildCentralAccountGlobalUnique
' Finds the central account that is already assigned to an employee or a user account in a target system.
#If Not SCRIPTDEBUGGER Then
	Imports System.Collections.Generic
	Imports System.Data
#End If

Public Overridable Function ccc_VI_AE_BuildCentralAccountGlobalUnique(ByVal uid_person As String, ByVal Lastname As String, ByVal Firstname As String) As String
            Dim i As Int32 = 0
            Dim account As String
            Dim accountPrefix As String
            Dim f As ISqlFormatter = Connection.SqlFormatter
			
			Dim LastnameLenLimit As Int32 = 11
			
			If String.IsNullOrEmpty(Lastname) AndAlso String.IsNullOrEmpty(Firstname) Then 
				Return String.Empty
			End If

            ccc_VI_AE_BuildCentralAccountGlobalUnique = String.Empty
			
			Firstname = ccc_Translit(Firstname).Substring(0, 1)
			Lastname = ccc_Translit(Lastname)
			If Lastname.Length > LastnameLenLimit Then
				Lastname = Lastname.Substring(0, LastnameLenLimit)
			End If

			account = Lastname

            accountPrefix = Firstname & account

            'fill existing addresses in a list 
            Dim existing As New List(Of String)
            Dim dummy As New Object()
            Dim dummyPerson As ISingleDbObject
            dummyPerson = Connection.CreateSingle("Person")
            Dim pattern As String = "%" & Lastname & "%"
            Dim myObjectKey As New DbObjectKey("Person", uid_person)
            Dim accountName As String
            Dim objectKeyString As String
            Dim objectKey As DbObjectKey

            Using rd As IDataReader = CType(dummyPerson.Custom.CallMethod("SearchCentralAccount", pattern), IDataReader)

                While rd.Read()

                    accountName = rd.GetString(rd.GetOrdinal("AccountName"))
                    objectKeyString = rd.GetString(rd.GetOrdinal("ObjectKeyPerson"))

                    If Not String.IsNullOrEmpty(objectKeyString) Then
                        objectKey = New DbObjectKey(objectKeyString)

                        'only addresses which not belong to the actual employee will be considered
                        If myObjectKey.Equals(objectKey) Then
                            Continue While
                        End If
                    End If

                    existing.Add(accountName)
                End While
            End Using

            While True
                Dim centralAccount As String = account

                ' Does not exists?
				If Not existing.Contains(centralAccount, StringComparer.InvariantCultureIgnoreCase) Then
					Return centralAccount.ToUpperInvariant()
				End If
				
				account = accountPrefix
				
				If i = 0
					' for next trial
					i = i + 1
					Continue  While
				End If
				
				' next trial
                account = accountPrefix & CStr(i)
				
				' adjustment
                i = i + 1

            End While
End Function