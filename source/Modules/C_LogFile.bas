Attribute VB_Name = "C_LogFile"
Option Explicit
Option Base 1
Public Logger As LogFile

Private Sub QuickIllustration()
    Dim mLogger As New LogFile        ' instantiate the log class as an object
    '--logger.CreateLogFile logFileName:="QuickIllustration" '<-- using legacy method
    mLogger.CreateLogFileByName "QuickIllustration"        ' or new alternative.
    mLogger.Log "Some string you want to log."
    mLogger.OpenLogFile
End Sub

'''
''' Example method demonstrating the the Log File class.
'''
Private Sub Example_03_LogFile_Demonstration()
    '#INCLUDE StartLogger
    Dim dirPath     As String
    Dim fileName     As String
    Dim iter        As Integer
    '// Instantiate the LogFile class.
    StartLogger        ' <-- only creates a new log file if none exists.
    '// Printing to the logfile uses enumerators for the output-formatting i.e. it is relatively straightforward.
    With Logger
        ' For example, we can define a title.
        .Log "Demonstration: ExampleLogger", LogTitle        'I prefer this style (it is more clean).
        ' <h1> Then we can add different headers like e.g.
        Call .Log("Start Example", LogHeader1, blankLineBeforeHeader:=False)        'but this style is equally okay.
        .Log "We will first list information about the current log file, and then print in a loop as an example."
        ' <h2> Logfile information
        .Log "Current LogFile Information", LogHeader2
        .Log "FilePath:= " & .FilePath, LogBullet1
        .Log "FileName:= " & .fileName, LogBullet2
        .Log "DirPath:= " & .DirectoryPath, LogBullet2
        ' <h2> Iteration loop as example
        .Log "Iteration Loop as an Example", LogHeader2
        For iter = 1 To 100
            If (iter Mod 10 = 0) Then
                .Log "At iteration:= " & iter & " value = " & VBA.rnd() * 100
            End If
        Next iter
        ' Now, let's clean-up and see what we got!
        .LogBlankLine
        .LogDividingLine
        .Log "How's it look? Now you try...", LogNoFormat
        .OpenLogFile
    End With        'logger
End Sub

'''
''' Displays the about string of the LogFile class.
'''
Private Sub Example_01_About()
    '#INCLUDE StartLogger
    StartLogger
    MsgBox "EXAMPLE ROUTINES FOR:" & VBA.vbCrLf & VBA.vbCrLf & Logger.About, vbInformation, Logger.Name
End Sub

'''
''' Initiates class variables if they do not exist.
'''
Private Sub StartLogger(Optional LogFileName As String = "0.LOG")
    Dim dirPath     As String
    Dim fileName_   As String
    '/* (0.) Instantiate the LogFile class.
    '        This is generally considered good coding practice i.e. instead of 'Dim logger as New LogFile' for every
    '        module, class and/or routine (like for the RegistryTRUxl class declaration above, but it can be done ;).
    '        This is especially true if you desired to declare a 'GLOBAL' (or module-level) instance of the LogFile
    '        class so that the same logfile can be used by multiple methods and/or functions.
    If Logger Is Nothing Then        ' no existing log file => create
        Set Logger = New LogFile
        '/* (1.) Before printing to the logfile you must first create the logfile
        '        if it does not already exist (e.g. from a previous method that
        '        already initiated the class object). If you wanted to 'continue'
        '        using a logfile that has already been initiated then this next
        '        statement should within the if statement above, but since this
        '        is an example we will start from scratch here.
        dirPath = Logger.RegDirectoryScratch        '<-- The LogFile class initializes using the dirPath defined in registry.
        '    This, of course, can be overwritten by defining a different dir path.
        '    If no definition exists, the LogFile class creates a folder named "scr"
        '    in the directory of the Excel file and defines the registry as such.
        fileName_ = LogFileName        '<-- A prefix is automatically generated for the FileName of the LogFile class,
        '    i.e. if this is left blank then a generic FileName is used.
        If Not Logger.CreateLogFile(dirPath, fileName_) Then
            Logger.AppErr.DisplayMessage
        Else
            Debug.Print "Log file successfully created :)"
            Debug.Print "|> FileName:= " & Logger.fileName
            Debug.Print "|> DirPath:= " & Logger.DirectoryPath
            Debug.Print "|> FilePath:= " & Logger.FilePath
        End If
    End If
End Sub

'''
''' Example that displays the different formatting styles/options of the Log File class.
'''
Private Sub Example_02_LogFile_Formatting_Options()
    '#INCLUDE StartLogger
    Dim dirPath         As String
    Dim fileName        As String
    Dim iter            As Integer
    '==========================================================================================================================
    ' Print to log file and display different formatting results.
    '==========================================================================================================================
    StartLogger        ' <-- only creates a new log file if none exists.
    '// LogFile class uses enums to handle formatting.
    With Logger
        .Log "supercalifragilisticexpialidocious :p"
        .LogBlankLine
        ' For example, we can define a title.
        .Log "|> logFormatType.logTitle writes as:", LogNoFormat
        .Log "Section Title", LogTitle, blankLineBeforeHeader:=False
        ' Heading
        .Log "|> logFormatType.logHeader1 writes as:", LogNoFormat
        .Log "Heading 1", LogHeader1, blankLineBeforeHeader:=False
        ' Sub-heading
        .Log "|> logFormatType.logHeader2 writes as:", LogNoFormat
        .Log "Heading 1.2", LogHeader2, blankLineBeforeHeader:=False
        ' Default output style:
        .Log "|> logFormatType.logDefault writes as:", LogNoFormat
        .Log "This is what the default output looks like.", LogDefault
        ' Line through.
        .LogBlankLine
        .Log "|> logFormatType.logLineThru writes as:", LogNoFormat
        .Log "This puts a line through the text.", LogLineThru
        ' No format.
        .LogBlankLine
        .Log "|> logFormatType.LogNoFormat writes as:", LogNoFormat
        .Log "You can also choose to output with no formatting (and no time stamp).", LogNoFormat
        ' Bullet points
        .LogBlankLine
        .Log "|> There are also different types of bullets written as:", LogNoFormat
        For iter = 1 To 6
            .Log "Bullet 1." & iter, LogBullet1
            .Log "Bullet 2." & iter, LogBullet2
        Next iter
        .LogBlankLine
        .Log "|> You can also log a dividing line to e.g. further separate sections."
        .LogDividingLine
        .Log "The LogFile class can be easily modified."
        .Log "Now you try...", LogNoFormat
        .OpenLogFile
    End With        ' logger
End Sub

'''
''' Changes the registry setting used to define the output directory.
'''
Sub Example_04_LogFile_Set_Output_Directory()
    '#INCLUDE StartLogger
    StartLogger
    ' The LogFile class defines ThisWorkbook.path & "\scr\" as the default directory and saves the path in the registry.
    ' This is automatically defined when the LogFile class is used for the first time (see the constructor method).
    ' This setting can be overwritten using the RegDirectoryScratch method.
    ' ** NOTE: only valid (existing) directory paths can be defined.
    Logger.RegDirectoryScratch = ThisWorkbook.Path & "\scr\"        ' <--- this updates a registry entry.
    If (Logger.AppErr.Number <> 0) Then
        Logger.AppErr.DisplayMessage
    Else
        Logger.OpenDirectory
    End If
End Sub

'''
''' Removes the registry setting defining the storage directory for the log files (for testing purposes).
''' This will not break anything, because the class will automatically define a default directory if none exists.
''' Sometimes an error is generated, but I cannot isolate it when stepping through i.e. it is a VBA quirk.
'''
Private Sub RemoveRegistrySetting()
    StartLogger
    Logger.RemoveRegistrySetting
End Sub

'''
''' Runs all examples.
'''
Sub RunAllExamples()
    '#INCLUDE Example_03_LogFile_Demonstration
    '#INCLUDE Example_01_About
    '#INCLUDE Example_02_LogFile_Formatting_Options
    Example_01_About
    Example_02_LogFile_Formatting_Options
    Example_03_LogFile_Demonstration
End Sub

