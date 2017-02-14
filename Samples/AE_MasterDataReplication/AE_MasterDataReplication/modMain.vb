Module modMain



    Sub Main()

        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty

        Try
            sFuncName = "Main"

            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Console.WriteLine("Calling GetSystemIntializeInfo() ", sFuncName)
           

        Catch ex As Exception
            Console.WriteLine("Completed with ERROR ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try

    End Sub

End Module
