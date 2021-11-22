'ccc_ExportDepartment2CSV
'Export Department data to CSV file (using OIM API) Created by S.Aksenov 16.06.2021
        Public Sub ccc_ExportDepartment2CSV(strFilePath As String)
            ' Enter script code here
            Dim StartTime As DateTime = Now()
            Dim strTimeStamp As String = VID_ISODatetimeForFilename(StartTime)

            Dim strCSVLine As String = ""

            Dim strLogFileName As String = String.Format("{0}\{1}_Department2CSV.log", strFilePath, strTimeStamp)
            Dim strFileName As String = String.Format("{0}\{1}_Department2CSV.csv", strFilePath, strTimeStamp)

            Dim strDelimiter As String = ";"

            ' Start exception handling block
            Try

                ' The code to observe comes here
                'Start Message --->
                VID_Write2Log(strLogFileName, String.Format("--> Script started at {0}", StartTime.ToString()))

                Dim colDepartments As IEntityCollection

                ' create the query object
                Dim qDepartment = Query.From("Department") _
                              .Select("UID_Department", "DepartmentName", "CustomProperty01")
                ' Load a collection of Person objects
                colDepartments = Session.Source.GetCollection(qDepartment)

                ' Number of objects to be affected --->
                VID_Write2Log(strLogFileName, String.Format("--> {0} department objects are going to be affected", colDepartments.Count().ToString()))

                ' Create a new StreamWriter object, second parameter = True --> append
                Using swFile As New System.IO.StreamWriter(strFileName, True)
                    ' Header information --->
                    swFile.WriteLine(String.Format("UID_Department{0}DepartmentName{0}DepartmentCustomProperty01", strDelimiter))

                    ' Run through the list
                    For Each colElement As IEntity In colDepartments
                        strCSVLine = String.Format("{1}{0}{2}{0}{3}",
                                                   strDelimiter,
                                                   colElement.GetValue("UID_Department"),
                                                   colElement.GetValue("DepartmentName"),
                                                   colElement.GetValue("CustomProperty01")
                                                   )
                        ' Write the line
                        swFile.WriteLine(strCSVLine)
                    Next

                End Using

                ' Script ended + execution time
                VID_Write2Log(strLogFileName, String.Format("--> Department export ended at {0}.{1}--> Execution time: {2} sec.", Now.ToString(), vbCrLf, (StartTime - Now).ToString("ss")))

            Catch ex As Exception

                ' If an exception occurs this code will be executed 
                VID_Write2Log(strLogFileName, ViException.ErrorString(ex))
                Throw New Exception("==> Debug: Department2CSV: ", ex)

                ' The execution continues after the End Try keyword
            End Try

        End Sub