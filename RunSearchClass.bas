Sub RunSearchClass()
    Application.ScreenUpdating = False

    'init
    Set instSearchClass = New SearchClass
    
    'run
    instSearchClass.MainSub
    
    'destory
    Set instSearchClass = Nothing
    
    MsgBox ("Search Complete")
    
    Application.ScreenUpdating = True
End Sub
