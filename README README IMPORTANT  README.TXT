
SIMPLE STEPS TO USING THE CLOCK

1) OPEN MakeDll  project first
   merely compile the project, as is as a dll ie (makedll.dll)

2 open project1
 
'''this sub....

Private Sub UserControl_Show()
    
' make sure that the makedll.dll is in the same path as project1
   Call loadDLL_webpage(WB1, App.Path & "\makedll.dll", 101)
 
End Sub

'and thats all there is too it.

What your basically doing is loading a webpage (the javascript code that
makes the clock) from a resource file which is compiled into the dll (makedll.dll)

The sub   Sub loadDLL_webpage  locates the webpage embedded in the dll and loads it
into the control which is just a webbrowser stretched across the usercontrol
any questions  email me at mrsnowlover@yahoo.com or IM on aol instant messenger
username mrsnowlover