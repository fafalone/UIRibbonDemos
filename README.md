# UIRibbonDemos
## Windows UI Ribbon Framework Demos

**Update: Crashing has been fixed.**

Windows applications frequently make use of the UI Ribbon in newer applications, including Explorer, Office, Wordpard, etc. While there's some controls implementing this kind of toolbar from scratch, and some exorbitantly expensive commercial controls that may or may not use the OS component ::cough:: CodeJock ::cough:: there's not been an implementation using the simple COM interfaces Microsoft applications use. This repository will show how to use those, with both simple and advanced demos.

### Requirements

1) Windows 7 or newer
2) You'll need either the Windows SDK v7.0 or newer, or to otherwise have obtained uicc.exe and rc.exe from it or Visual Studio.
3) An up to date twinBASIC; not sure where the cutoff would be but always good to have a more recent release.
4) Becoming familiar with the XAML-based [Ribbon Markup Language](https://learn.microsoft.com/en-us/windows/win32/windowsribbon/windowsribbon-schema) to create the .xml files describing controls and commands. A good example for learning the syntax [also accompanies this demo](https://www.codeproject.com/Articles/160542/Windows-7-Goodies-in-C-Introduction-to-the-Ribbon), although the example itself is C++. 

### First demo - Hello Ribbon!

![Screeshot1](https://i.imgur.com/Ns39N3J.jpg)

For our first ribbon application, we'll use this simple one, based on a [pure C version by Stephen Wiria](https://www.codeproject.com/Articles/119319/Windows-Ribbon-Framework-in-Win32-C-Application).

We'll start from the xml:

#### Preparing the project files
Once you have the ribbon.xml file and the \Res folder containing the bitmap images for your controls, you can proceed to preparing the project.
1) Use uicc.exe to compile the XML. This is easiest if you have Visual Studio command prompt available, but you can substitute full paths or drop uicc.exe in the ribbon folder. We want not just the compiled file, but we want uicc to prepare a resource file containing all the strings and bitmaps correctly named so importation into twinBASIC is nice and simple. For this we use the following command:
    
   `uicc.exe ribbon.xml ribbon.bml /res:ribbon.rc`

   The first file is the compiled binary file; while you can use that for manual importation, it's already copied into ribbon.rc, so you don't need to worry about that for our method.

2) Compile ribbon.rc with rc.exe -- this is simple, in the same prompt, just use `rc ribbon.rc`. This will produce a .res file, which you might already be familiar with as this is the format we use in VB6 for resource files.

#### Import into twinBASIC and set up project
1) twinBASIC does not currently support importing .res files directly-- but it does as part of the .vbp import process. This repository contains ImportRibbon.vbp, an otherwise empty VBP file that will trigger twinBASIC to import the ribbon.res file in the same directory as the .vbp. Open twinBASIC, from the new project tab select 'Import from VBP...' and choose our ImportRibbon.vbp. This will fill the resources folder with our binary UIFILE from ribbon.bml, a BITMAP folder containing all our images, and a string table containing all the control captions etc.
2) You'll want to add a Form, and a class named clsRibbonEvents. Then open up the Settings, set the name, and anything else you want, and go down to `COM Type Library / Active-X References`, click the TWINPACK PACKAGES button, and add a reference to `twinBASIC Shell Library v4.13.175` (or the latest version).
3) Save the project, and you're now ready to code, which is actually simpler than everything we've done so far.

#### Code for the ribbon
1) The form sets everything up, then the class handles the events the ribbon raises to let us know about command clicks and other information.
2) The Form code declares a variable for the UI Ribbon Framework coclass, the events class, and a handler for the command click it raises:
   ```
    Private pFramework As UIRibbonFramework
    Private WithEvents pUIApp As clsRibbonEvents
    
    Private Sub Form_Load() Handles Form.Load
        Set pFramework = New UIRibbonFramework
        Set pUIApp = New clsRibbonEvents
        pFramework.Initialize Me.hWnd, pUIApp
        pFramework.LoadUI GetModuleHandleW(), StrPtr("APPLICATION_RIBBON")
    End Sub
    
    Private Sub Form_Terminate() Handles Form.Terminate
        pFramework.Destroy
        Set pFramework = Nothing
        Set pUIApp = Nothing
    End Sub
    
    Private Sub pUIApp_OnRibbonCmdExecute(ByVal commandId As Long, ByVal verb As UI_EXECUTIONVERB, ByVal key As LongPtr, currentValue As Variant, ByVal commandExecutionProperties As IUISimplePropertySet, returnValue As Long) Handles pUIApp.OnRibbonCmdExecute
        List1.AddItem "You clicked: CommandId=" & commandId & ", Verb=" & verb
    End Sub
   ```
   
   All of those interfaces and the GetModuleHandle API are already declared in tbShellLib; that's the entirety of the form code. In the class, we have:
   ```
        Private Sub IUICommandHandler_Execute(ByVal commandId As Long, ByVal verb As UI_EXECUTIONVERB, key As PROPERTYKEY, currentValue As Variant, ByVal commandExecutionProperties As IUISimplePropertySet) Implements IUICommandHandler.Execute
            Dim hr As Long
            Dim pv As Variant
            If VarPtr(currentValue) <> 0 Then
                VariantCopy pv, currentValue
            End If
            RaiseEvent OnRibbonCmdExecute(commandId, verb, VarPtr(key), pv, commandExecutionProperties, hr)
            Err.ReturnHResult = hr
        End Sub
        
        Private Sub IUICommandHandler_UpdateProperty(ByVal commandId As Long, key As PROPERTYKEY, currentValue As Variant, newValue As Variant) Implements IUICommandHandler.UpdateProperty
            Dim hr As Long
            Dim pv As Variant
            Dim pnv As Variant
            Dim bValid As Boolean
            If VarPtr(currentValue) <> 0 Then
                VariantCopy pv, currentValue
            End If
            RaiseEvent OnRibbonUpdateProperty(commandId, VarPtr(key), pv, pnv, bValid, hr)
            If bValid Then
                VariantCopy newValue, pnv
            End If
            Err.ReturnHResult = hr
        End Sub
   ```

3) Compile and run, that's all there is to it! **NOTE:** This will not run from the IDE, since when running from the IDE, like in VB6, it uses the resources in twinBASIC, not our resources. Since the object extracts the resources itself from the handle to the exe, it won't be able to find the ribbon resources.
   
**The end result, the completed project containing the results of these steps, is in UIRibbonDemo.twinproj.** The rest of the files, including intermediates, are also provided. 

---

I'll be adding more advanced demos shortly!
