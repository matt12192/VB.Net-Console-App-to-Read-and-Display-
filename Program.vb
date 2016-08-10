Module Program

    'Program to read in details of Atheles and distance of their throws
    'The Program should select the furthest throw and declare that athelete as the winner
    Dim AthletesName(19) As String
    Dim AthletesNumber(19) As String
    
    Dim ThrowDistance(19, 3) As String

    Dim bThrowerCount As Byte
    Dim bytPosition As Byte = 0
    Dim shFurthestThrow As Short = -99
    Dim strMessage As String = ""

    Sub Main()

        Initialise()
        Menu()

    End Sub


    Sub Menu()
        'Display the program menu
        Dim bytChoice As Byte = 0

        Do
            Try
                System.Console.Clear()
                System.Console.WriteLine(vbLf & "Mens Javelin Tournament")
                System.Console.WriteLine()
                System.Console.WriteLine()
                System.Console.WriteLine(vbLf & "1. Enter Athletes Details and their Throw Distance")
                System.Console.WriteLine("2. Display Throw Results")
                System.Console.WriteLine("3. Display Overall Winner")
                System.Console.WriteLine("4. Clear all readings")
                System.Console.WriteLine("5. Exit")
                System.Console.Write(vbLf & vbLf & "Please enter choice : ")
                bytChoice = System.Console.ReadLine

                Select Case bytChoice
                    Case 1
                        EnterAthletesDetails()
                    Case 2
                        DisplayThrowResults()
                    Case 3
                        DisplayOverallWinner()
                    Case 4
                        ClearAllReadings()
                    Case 5
                        ExitSystem()
                    Case Else
                        System.Console.WriteLine(vbLf & "You have entered an invalid choice")
                        PressEnter()
                End Select
            Catch ex As Exception
                System.Console.WriteLine(vbLf & "You have entered an invalid characters")
                PressEnter()
            End Try
        Loop While (bytChoice <> 5)


    End Sub

    Sub Initialise()

        Dim bCount As Byte = 0
        For bCount = 0 To 19
            AthletesName(bCount) = ""
            AthletesNumber(bCount) = ""
            ThrowDistance(bCount, 0) = -999
            ThrowDistance(bCount, 1) = -999
            ThrowDistance(bCount, 2) = -999
            ThrowDistance(bCount, 3) = -999
        Next
        strMessage = " "
        Display()

    End Sub
    Sub Display()

        Console.BackgroundColor = ConsoleColor.DarkGreen
        Console.ForegroundColor = ConsoleColor.White
        Console.Clear()
        Console.Title = "Mens Javelin Tournament"

    End Sub

    Sub EnterAthletesDetails()
        'Allows up to 20 names and number of Atheletes and their throws to be entered
        'The user can stop entering the information of Athletes by typing in -999 when asked for the number of the next athlete
        Dim intAthletesNumber As Integer = 0
        Dim strThrowerName As String = ""
        Dim bThrowerCount As Byte = 0
        Dim bCounter As Byte = 0 'Keep track of throws
        System.Console.Clear()

        Do
            System.Console.WriteLine(vbLf & vbLf & "Mens Javelin Tournament")
            System.Console.WriteLine(vbLf & "Enter -999 to return to menu: ")
            System.Console.WriteLine()
            System.Console.Write(vbLf & "Enter Number [101 - 199] : ")
            intAthletesNumber = System.Console.ReadLine
            If intAthletesNumber > -999 Then

                AthletesNumber(bThrowerCount) = ValidateNumber(intAthletesNumber, 100, 200)
                System.Console.Write(vbLf & "Enter Name : ")
                strThrowerName = System.Console.ReadLine
                AthletesName(bThrowerCount) = strThrowerName


                'Code to read in throws
                EnterThrows(bThrowerCount)

                bThrowerCount = bThrowerCount + 1
            Else
                Exit Sub
            End If
            System.Console.Clear()

        Loop While (strThrowerName <> "#") And (bThrowerCount <> 19)
        PressEnter()

    End Sub

    Sub EnterThrows(ByVal bThrowerCount As Byte)
        'This code will read in the Throw results for each athlete
        Dim sDistance As Single = 0
        Dim sFurthestDistance As Single = -99
        System.Console.Write(vbLf & "Enter Throw 1: ")

        sDistance = System.Console.ReadLine
        'code to check for a Throw fault
        If sDistance = -1 Then
            System.Console.WriteLine("Invalid Throw please Throw again!")
            System.Console.Write(vbLf & "Re-Enter Throw 1: ")
            sDistance = System.Console.ReadLine

        End If

        'Code to check for a fault

        ThrowDistance(bThrowerCount, 0) = ValidateNumber(sDistance, 1, 100)
        'Code to check range
        System.Console.Write(vbLf & "Enter Throw 2: ")
        sDistance = System.Console.ReadLine
        'code to check for a Thriw fault
        If sDistance = -1 Then
            System.Console.WriteLine("Invalid Throw please Throw again!")
            System.Console.Write(vbLf & "Re-Enter Throw 2: ")
            sDistance = System.Console.ReadLine

        End If

        ThrowDistance(bThrowerCount, 1) = ValidateNumber(sDistance, 1, 100)

        System.Console.Write(vbLf & "Enter Throw 3: ")
        'Code to check for a fault

        sDistance = System.Console.ReadLine
        'code to check for a Throw fault using
        If sDistance = -1 Then
            System.Console.WriteLine("Invalid Throw please Throw again!")
            System.Console.Write(vbLf & "Re-Enter Throw 3: ")
            sDistance = System.Console.ReadLine

        End If

        ThrowDistance(bThrowerCount, 2) = ValidateNumber(sDistance, 1, 100)

        'Select best throw

        If ThrowDistance(bThrowerCount, 0) > sFurthestDistance And ThrowDistance(bThrowerCount, 0) > -99 Then
            sFurthestDistance = ThrowDistance(bThrowerCount, 0)
        End If

        If ThrowDistance(bThrowerCount, 1) > sFurthestDistance And ThrowDistance(bThrowerCount, 0) > -99 Then
            sFurthestDistance = ThrowDistance(bThrowerCount, 1)
        End If

        If ThrowDistance(bThrowerCount, 2) > sFurthestDistance And ThrowDistance(bThrowerCount, 0) > -99 Then
            sFurthestDistance = ThrowDistance(bThrowerCount, 2)
        End If
        ThrowDistance(bThrowerCount, 3) = sFurthestDistance

    End Sub

    Function ValidateNumber(ByVal number, ByVal min, ByVal max)
        'Code to check range
        If number < min Or number > max Then
            System.Console.WriteLine("Please check the value and enter a number between" & min & " and " & max)
            number = 0
            System.Console.Write(vbLf & "Please re-enter the correct value ")
            number = System.Console.ReadLine

            Return number
        Else
            Return number
            Exit Function

        End If

    End Function

    Sub DisplayThrowResults()
        Dim bThrowerCount As Byte = 0
        System.Console.Clear()
        System.Console.WriteLine(vbLf & vbLf & "Mens Javelin Tournament ")
        System.Console.WriteLine()
        System.Console.WriteLine("Name " & vbTab & "Number " & vbTab & "Throw 1 " & vbTab & "Throw 2 " & vbTab & "Throw 3 " & vbTab & "Best Throw ")
        System.Console.WriteLine()
        For bThrowerCount = 0 To 19
            If AthletesName(bThrowerCount) > "" Then
                System.Console.Write(AthletesName(bThrowerCount) & vbTab & AthletesNumber(bThrowerCount))
                System.Console.WriteLine(vbTab & ThrowDistance(bThrowerCount, 0) & vbTab & vbTab & ThrowDistance(bThrowerCount, 1) & vbTab & vbTab & ThrowDistance(bThrowerCount, 2) & vbTab & vbTab & ThrowDistance(bThrowerCount, 3))

            End If
        Next
        PressEnter()
    End Sub

    Function FurthestThrow() As String

        Dim bCount As Byte = 0

        For bCount = 0 To 19
            If (ThrowDistance(bCount, 3) > shFurthestThrow) And (ThrowDistance(bCount, 3) <> -999) Then
                shFurthestThrow = ThrowDistance(bCount, 3)
                bytPosition = bCount
            End If
        Next
        strMessage = "The Winner is " & AthletesNumber(bytPosition) & "  " & AthletesName(bytPosition) & " With a Throw distance of " & shFurthestThrow & " Metres!"
        Return strMessage
    End Function

    Sub DisplayOverallWinner()
        'Display the details of overall winner

        System.Console.Clear()
        System.Console.WriteLine(vbLf & vbLf & "Mens Javelin Tournament ")
        System.Console.WriteLine()

        FurthestThrow()
        System.Console.WriteLine(vbLf & strMessage)
        Console.WriteLine(vbLf & "Congratulations you have won a prize")
        PressEnter()
    End Sub

    Sub ClearAllReadings()
        'Ask user to confirm clear
        Dim strClearResponse As String = ""
        System.Console.Write(vbLf & vbLf & "Are you sure you want to clear all details?")

        strClearResponse = System.Console.ReadLine
        If strClearResponse = "Yes" Or strClearResponse = "yes" Or strClearResponse = "YES" Then
            Initialise()
            System.Console.WriteLine(vbLf & "ALL Atheltes details have been cleared")
            PressEnter()
        End If
    End Sub

    Sub PressEnter()
        'Pause while the user presses enter key
        System.Console.Write(vbLf & vbLf & "Press [Enter] to continue")
        System.Console.ReadLine()

    End Sub

    Sub ExitSystem()
        'Ask user to confirm clear
        Dim strClearResponse As String = ""
        System.Console.Write(vbLf & vbLf & "Are you sure you want to exit the Javelin program? [Yes/No]")
        strClearResponse = System.Console.ReadLine
        If strClearResponse = "Yes" Or strClearResponse = "yes" Or strClearResponse = "YES" Or strClearResponse = "Y" Or strClearResponse = "Yes" Then
            Initialise()
            System.Console.WriteLine(vbLf & "Thank you for using the Mens Javelin Tournament Program ")
            PressEnter()
            End
        Else
            Console.Clear()
            Menu()
        End If

    End Sub

End Module
