Attribute VB_Name = "Module2"
Dim ini(1000) As String
Dim maxcnt As Long
Sub writeini(lpFilename, lpApplicationName, lpKeyName, lpString)
    'This function save data to the ini file
    
    'Checks  to see if ini file exists
    If Dir(lpFilename) = "" Then
        'Open and creates a new file
        Open lpFilename For Output As #1
        Close #1
    End If
        
    
    'First Load the ini into memory
    LoadInitoMemory (lpFilename)
    
    'If maxcnt = 0 Then Exit Sub
    'Now check to see if application exists
    'If -1 then its not found
    appnum = ApplicationExists(lpApplicationName)
'Stop
    'Looking for KeyName
    'If -1 then its not found
    keynum = KeyNameExists(lpApplicationName, lpKeyName)
'Stop
    'If ApplicationName is there
    If Not appnum = -1 Then
        'If keyName is there
        If Not keynum = -1 Then
            ini(keynum) = lpKeyName & "=" & lpString
        Else

            pos = FindAvailableApplicationEntry(appnum)
            pos = InsertEntry(pos)
            ini(pos) = lpKeyName & "=" & lpString
        End If
    Else
        
        'If ApplicationName doesn't exists then add to end
            'This Part creates a new entry
            maxcnt = maxcnt + 1
            ini(maxcnt) = "[" & lpApplicationName & "]"
            maxcnt = maxcnt + 1
            ini(maxcnt) = lpKeyName & "=" & lpString
            maxcnt = maxcnt + 1
            ini(maxcnt) = ""
    End If
    
    
    
    fnum = FreeFile
    Open lpFilename For Output As #1
        For lop = 0 To maxcnt - 1
         Print #fnum, ini(lop)
        Next lop
    Close #1
'Stop
End Sub


Function ApplicationExists(lpApplicationName) As Long
    
     Do Until cnt = maxcnt
        
        info = ini(cnt)
        'Is it the start of the appName?
        If Left$(info, 1) = "[" Then
            'Extracts the appname from the line
            Appname = LCase$(Mid$(info, 2, Len(info) - 2))
              'If it's found the correct 1
              If LCase$(lpApplicationName) = Appname Then
'Stop
                  ApplicationExists = cnt
                  Exit Function
              End If
        End If
        cnt = cnt + 1
     Loop
     
    ApplicationExists = -1

End Function

Function KeyNameExists(lpApplicationName, lpKeyName) As Long
   'Stop
     Do Until cnt = maxcnt
        info = ini(cnt)
       'AppName matches then scan entried below it
      If Not info = "" Then
       If LCase$(Mid$(info, 2, Len(info) - 2)) = LCase$(lpApplicationName) Then
'        Stop
        Do
        'Get next line
        cnt = cnt + 1
        info = ini(cnt)
         
         appstart = Left$(info, 1)           'Gets letter for the [
         If appstart = "[" Or appstart = "" Then
            KeyNameExists = -1
            Exit Function       'If [ accurs then get out
        End If
         If LCase$(Left$(info, Len(lpKeyName))) = LCase$(lpKeyName) Then
            'Match Made
            KeyNameExists = cnt
            'Close #fnum
            Exit Function
         End If
            'KeyNameExists = False
            'Close #fnum
            'Exit Function
         Loop
         End If
        
        End If
        cnt = cnt + 1
       Loop
       KeyNameExists = -1
       

End Function

Sub LoadInitoMemory(filename)
    'This sub loads an ini file into an array for easy modifying
    fnum = FreeFile
    Open filename For Input As fnum
    Do Until EOF(fnum)
    Line Input #fnum, info
    ini(cnt) = info
    cnt = cnt + 1
    Loop
    Close fnum
    maxcnt = cnt
'Stop

End Sub

Function InsertEntry(lpPosition)
    'This function will move all the array entries and return the position
    'of where to insert the new value
    maxcnt = maxcnt + 1
    For lop = maxcnt To lpPosition Step -1
     ini(lop) = ini(lop - 1)
    Next lop
     
    InsertEntry = lpPosition
'Stop
End Function

Function FindAvailableApplicationEntry(lpPosition)
    'This finds the last possible entry space in the AppName
    For lop = lpPosition To maxcnt
        If Left(ini(lop + 1), 1) = "[" Or lop = maxcnt Then
            FindAvailableApplicationEntry = lop
            Exit For
        End If
    Next lop
'Stop
            
End Function

Function Readini(lpFilename, lpApplicationName, lpKeyName) As Variant
    'This read information from an ini
    'It uses Variant to decide if its a string or number
    
    'Loads the ini to memory b4 startin workin
    LoadInitoMemory (lpFilename)
    
    'Searches for the appname
    appnum = ApplicationExists(lpApplicationName)
    
    'If Application doesnt exit then exit out with error
    If appnum = -1 Then
        'Readini = "Error - Application Name does not exist"
        Readini = -1
        Exit Function
    End If
    
    'It now looks for the keyname in the appname
    keynum = KeyNameExists(lpApplicationName, lpKeyName)

    'Error out if key isn't found
    If keynum = -1 Then
        'Readini = "Error - The Keyname does not exist in the specified ApplicationName"
        Readini = -1
        Exit Function
    End If
    
    'If the code gets here, then it means all is well :o)
    Readini = Mid$(ini(keynum), Len(lpKeyName) + 2, Len(ini(keynum)))
        
End Function
