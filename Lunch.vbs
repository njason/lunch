'Selects where to go to Lunch for people who are too lazy to decide.
'Jason Biegel

Randomize

Const DEBUG_     = False    'displays debug information (suggested to be used in command prompt with wscript) 
Const LOCAL_     = False    'displays a message on the computer of the results
Const WEBSITE_     = True 'displays results to a website on this machine's local IIS server, you need IIS support for this to work 
Const PICS_    = True 'an extra gag for the website, randomly selects a picture off the internet and displays it, will do nothing if website is not enabled

Const numPics = 70    'the amount of pictures to choose from 

Const ForReading = 1
Const ForWriting = 2
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim Places()
Dim Weight()
Dim Rank()
Dim Choices()
Dim ThisWeek()

'INPUT 
'reads in the places to be chosen for lunch this week
Set objFile = objFSO.OpenTextFile("LunchPlaces.txt", ForReading)
Index = 0
Do Until objFile.AtEndOfStream
    Place = objFile.ReadLine
    Taken = False
    ReDim Preserve Places(Index)
    Places(Index) = Place
    Index = Index + 1
Loop
objFile.Close

'reads in the amount of times each place has been chosen
ReDim Weight(Index-1) 
Set objFile = objFSO.OpenTextFile("LunchHistory.txt", ForReading)
Do Until objFile.AtEndOfStream
    Place = objFile.ReadLine
    For X = 0 To Index-1
        If Place = Places(X) Then
            Weight(X) = objFile.ReadLine
            If Weight(X) = "" Then
                Weight(X) = 0
            End If
            Stop
        End If
    Next
Loop
objFile.Close

'reads in places already eaten at this week 
NumTaken = 0

If WeekDay(Date) <> vbMonday Then
    Set objFile = objFSO.OpenTextFile("ThisWeek.txt", ForReading)
    Do Until objFile.AtEndOfStream
        ReDim Preserve ThisWeek(NumTaken) 
        ThisWeek(NumTaken) = objFile.ReadLine
        NumTaken = NumTaken + 1
    Loop
    objFile.Close
    If NumTaken = Index Then
        NumTaken = -1
    End If
End If
'INPUT - END

'SELECTION
'assigns rank to each place determined by how many times we have gone there
ReDim Rank(Index-1)
For X = 0 To Index-1
    Rank(X) = Int(Weight(X) * (5 * Rnd + 1))
Next

'determines the candidates for being chosen for lucnh 
'(which ever has the lowest rank)
Min = -1
Index2 = 0
For X = 0 To Index-1
    Taken = False
    For Y = 0 To NumTaken-1
        If ThisWeek(Y) = Places(X) Then
            Taken = True
            Stop 
        End If
    Next
    If Taken = False Then
        If Min > Rank(X) Or Min = -1 Then
            Min = Rank(X)
            Index2 = 0
            ReDim Choices(Index2)
            Choices(Index2) = Places(X) 
        ElseIf Min = Rank(X) Then
            Index2 = Index2 + 1
            ReDim Preserve Choices(Index2)
            Choices(Index2) = Places(X)
        End If
    End If
Next
'SELECTION - END 

'DEBUGGING
If DEBUG_ = True Then
    Wscript.Echo
    Wscript.Echo("The Places, Weight, and Rank:")
    For X = 0 To Index-1
        Wscript.Echo(Places(X) & ", " & Weight(X) & ", " & Rank(X)) 
    Next
    Wscript.Echo
    Wscript.Echo("Taken this week:")
    For X = 0 To NumTaken-1
        Wscript.Echo(ThisWeek(X))
    Next
    Wscript.Echo
    Wscript.Echo("Choices:") 
    For X = 0 To Index2
        Wscript.Echo(Choices(X))
    Next
    Wscript.Echo
End If
'DEBUGGING - END

'PICTURE
If WEBSITE_ = True And PICS_ = True Then
    pick = Int(rnd * numPics) 
    Set objFile = objFSO.OpenTextFile("pics.txt", ForReading)
    picCount = 0
    Do Until objFile.AtEndOfStream
        temp = objFile.ReadLine
        If picCount = pick Then
            pic = temp 
        End If
        picCount = picCount + 1
    Loop
End If
'PICTURE - END

'OUTPUT
'updates the count on the chosen place
Choice = Int(Index2+1 * Rnd)
For X = 0 To Index-1
    If Choices(Choice) = Places(X) Then
        Weight(X) = Weight(X) + 1
        Stop
    End If
Next

'writes the history of each place
Set objFile = objFSO.CreateTextFile("LunchHistory.txt ")
For X = 0 To Index-1
    objFile.WriteLine(Places(X))
    objFile.WriteLine(Weight(X))
Next
objFile.Close

'updates the places that were already chosen
Set objFile = objFSO.CreateTextFile ("ThisWeek.txt")
For X = 0 To NumTaken-1
    objFile.WriteLine(ThisWeek(X))
Next
objFile.WriteLine(Choices(Choice))
objFile.Close

If WEBSITE_ = True Then
    Set objFile = objFSO.CreateTextFile ("C:\Inetpub\wwwroot\Lunch\today.html")
    objFile.WriteLine("<html>")
    objFile.WriteLine("    <head>")
    objFile.WriteLine("        <title>")
     objFile.WriteLine("            Today's lunch for the Co-ops at Marsh")
    objFile.WriteLine("        </title>")
    objFile.WriteLine("    </head>")
    objFile.WriteLine ("    <body>")
    objFile.WriteLine("        <h3>")
    objFile.WriteLine("            Lunch will be at " & Choices(Choice) & ".")
    objFile.WriteLine ("        </h3>")
    objFile.WriteLine("        " & WeekdayName(Weekday(Date)) & ", " & Month(Date) & "/" & Day(Date) & "/" & Year(Date)) 
    If PICS_ = True Then
        objFile.WriteLine("        <br /> <br /> <br />")
        objFile.WriteLine("        Motivational statement of the day <br />")
         objFile.WriteLine("        <img src=""" & pic & """>")
    End If
    objFile.WriteLine("    </body>")
    objFile.WriteLine("</html>")     
    objFile.Close
End If

If LOCAL_ = True Then
    Wscript.Echo("Lunch will be at " & Choices(Choice) & ".")
End If

'OUTPUT - END

==================================================================================== 

'Selects where to go to Lunch for people who are too lazy to decide like me.
'Jason Biegel
Dim Places()
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("LunchPlaces.txt", ForReading)
Index = 0
Do Until objFile.AtEndOfStream
    ReDim Preserve Places(Index)
    Places(Index) = objFile.ReadLine
    Index = Index + 1 
Loop
objFile.Close
randomize
rnum = int(Index * rnd)
wscript.Echo("Lunch will be at " & Places(rnum) & ".")
