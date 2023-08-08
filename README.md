# UIRibbonDemos
## Windows UI Ribbon Framework Demos

**Update (Augest 7th): Gallery Intro Demo Released!** The finished project is UIRibbonDemoSGallery.twinproj,and all the intermediates are in \UIRibbonDemoSGallery. More info down at the bottom after the Intermediate demo. **Update:** Same day minor bug fix, custom border size (1-15) now works.


**Update (August 1st): Intermediate Demo updated.**\
Bug fix: Toggle button for context tabs wasn't working. \
Bug fix: Shield icon wasn't replaced with app icon when running from IDE.\
I have not been able to get Recent Items pinning working correctly... this version implements a method I found that should supposedly read the changes to pinned status, but it's returning false always for every item. There is code that gets an updated pin status if you click the item. I've left all my debug code for this in, and activated the IUIEventLogger, if anyone wantsto play around with this. Will keep working on it but didn't want to delay the other fixes.

**Update (July 30th): Intermediate Demo updated!** Since the Recent Items category was there, I thought filling it shouldn't wait for the advanced demo. Details at end of readme.

**Update (July 28th): The Intermediate Demo is now available! Scroll down to check it out!** The finished project is UIRibbonDemoIntermediateA.twinproj, and the intermediates are the ones with an I suffix.

**Update: Crashing has been fixed.**

Windows applications frequently make use of the UI Ribbon in newer applications, including Explorer, Office, Wordpard, etc. While there's some controls implementing this kind of toolbar from scratch, and some exorbitantly expensive commercial controls that may or may not use the OS component ::cough:: CodeJock ::cough:: there's not been an implementation using the simple COM interfaces Microsoft applications use. This repository will show how to use those, with both simple and advanced demos.

### Requirements

1) Windows 7 or newer
2) You'll need either the Windows SDK v7.0 or newer, or to otherwise have obtained uicc.exe and rc.exe from it or Visual Studio.
3) An up to date twinBASIC; not sure where the cutoff would be but always good to have a more recent release.
4) Becoming familiar with the XML-based [Ribbon Markup Language](https://learn.microsoft.com/en-us/windows/win32/windowsribbon/windowsribbon-schema) to create the .xml files describing controls and commands. A good example for learning the syntax [also accompanies this demo](https://www.codeproject.com/Articles/160542/Windows-7-Goodies-in-C-Introduction-to-the-Ribbon), although the example itself is C++.

> [!NOTE]
> I recommend the Ribbon Designer in the Delphi Ribbon Frame by JAM-Software. [an open source project here on GitHub](https://github.com/JAM-Software/RibbonFramework). While it's not written in tB or VB6, it can be compiled without issue from source if you don't want to download the binary from the free Delphi IDE. It's a GUI-based designer that greatly simplifies the process of generating the XML, although you will still want to familiaring yourself with it, since the tool doesn't explain how it all works.

### First demo - Hello Ribbon!

![Screeshot1](https://i.imgur.com/Ns39N3J.jpg)

For our first ribbon application, we'll use this simple one, based on a [pure C version by Stephen Wiria](https://www.codeproject.com/Articles/119319/Windows-Ribbon-Framework-in-Win32-C-Application).

The final .twinproj for this is UIRibbDemo.twinproj.

We'll start from the xml:

#### Preparing the project files
Once you have the ribbon.xml file and the \Res folder containing the bitmap images for your controls, you can proceed to preparing the project.
1) Use uicc.exe to compile the XML. This is easiest if you have Visual Studio command prompt available, but you can substitute full paths or drop uicc.exe in the ribbon folder. We want not just the compiled file, but we want uicc to prepare a resource file containing all the strings and bitmaps correctly named so importation into twinBASIC is nice and simple. For this we use the following command:
    
   `uicc.exe ribbon.xml ribbon.bml /res:ribbon.rc`/header:ribbon.h

   The first file is the compiled binary file; while you can use that for manual importation, it's already copied into ribbon.rc, so you don't need to worry about that for our method.

2) Compile ribbon.rc with rc.exe -- this is simple, in the same prompt, just use `rc ribbon.rc`. This will produce a .res file, which you might already be familiar with as this is the format we use in VB6 for resource files.

#### Import into twinBASIC and set up project
1) twinBASIC does not currently support importing .res files directly-- but it does as part of the .vbp import process. This repository contains ImportRibbon.vbp, an otherwise empty VBP file that will trigger twinBASIC to import the ribbon.res file in the same directory as the .vbp. Open twinBASIC, from the new project tab select 'Import from VBP...' and choose our ImportRibbon.vbp. This will fill the resources folder with our binary UIFILE from ribbon.bml, a BITMAP folder containing all our images, and a string table containing all the control captions etc.
2) You'll want to add a Form, and a class named clsRibbonEvents. Then open up the Settings, set the name, and anything else you want, and go down to `COM Type Library / Active-X References`, click the TWINPACK PACKAGES button, and add a reference to `twinBASIC Shell Library v4.13.175` (or the latest version).
3) Save the project, and you're now ready to code, which is actually simpler than everything we've done so far.

#### The Basic Demo
1) The form sets everything up, then the class handles the events the ribbon raises to let us know about command clicks and other information.
2) The Form code declares a variable for the UI Ribbon Framework coclass, the events class, and a handler for the command click it raises:
   ```vb6
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
   ```vb6
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

3) Compile and run, that's all there is to it! **NOTE:** This can run from the IDE, but must be compiled first in order to generate a resource containing binary that can be loaded. You'll need to do this before every run from the IDE, regardless of whether you've made any changes.
   
**The end result, the completed project containing the results of these steps, is in UIRibbonDemo.twinproj.** The rest of the files, including intermediates, are also provided. 

---

## GUI-based Ribbon Designer
Before getting into the Intermediate Demo, it's worth noting I found a GUI-based designer that simplifies the XML generation by quite a bit. It's open source, but written in Delphi. You can compile it from the free IDE. It's contained in [JAM-Software's RibbonFramework for Delphi](https://github.com/JAM-Software/RibbonFramework). It will save the XML, but the Build feature isn't helpful for us. You'll still need to know the concepts below.

## The Intermediate Demo

Now that we've estashlished how to get your application to show a Ribbon and the basic functionality, let's dig into some of the features!

### Font Control
The highlight of this demo is showing how to link up the Font Control and some additional buttons to a RichEdit control:

![FontControl](https://i.imgur.com/qHyhR3O.gif)

This is by far the most complex control in the Ribbon. It displays a full set of font font options like you see in Wordpad, including an automatically populated dropdown showing each font rendered in it's own face. This is tied to a RichEdit control, and  you can use it to set the font of the current selection. If you move the caret in the RichEdit control, the font displayed in the Ribbon is updated to that of the current selection.\
You'll find a lot of useful code for this, including dealing with true NULL Variants that crash tB when improperly handled, reading and setting values in an IPropertyStore, and converting back and forth between two types normally unsupported in VB and which cause errors when not specially handled: Variants of type VT_UI4 (unsigned long), and VT_LPWSTR, a string type similar to but different in important ways from VB/tB native strings (VT_BSTR). Finally, we get a chance to use tB's new Decimal type! VB6 allowed some use of these in a Variant, but  without the real native type, it would require difficult manual handling for our purposes.

It's placed in the control with a Command entry and View entry:

```
    <Command Name="cmdRichFont" Id="1779">
      <Command.Keytip>
        <String Id="1780">F</String>
      </Command.Keytip>
    </Command>
    <Group CommandName="cmdGroupRichFont" SizeDefinition="OneFontControl">
        <FontControl CommandName="cmdRichFont" FontType="RichFont"/>
    </Group>
```

`FontType` here refers to the version of the font control you want; this one has the most options, and there's a basic and medium version too. Note that the buttons and other controls on it all come up automatically; you don't need to supply your own images or commands.

### Dropdown Buttons
On the Design context tab, you'll find an Edit button with an arrow indicating a dropdown. This button displays a dropdown menu of basic edit commands. You'll note these same commands are re-used in a couple other places; you don't need to define different ones for e.g. the context menu or other tab.

![image](https://github.com/fafalone/UIRibbonDemos/assets/7834493/c8a5da92-0b8c-45fa-910b-6238eab9db11)

Thankfully, this one is easy:

```
            <Group CommandName="cmdGroup6" SizeDefinition="OneButton">
              <DropDownButton CommandName="cmdDropDownButton">
                <MenuGroup>
                  <Button CommandName="cmdCut"/>
                  <Button CommandName="cmdCopy"/>
                  <Button CommandName="cmdPaste"/>
                </MenuGroup>
              </DropDownButton>
            </Group>
```

### SplitButtons:
The original demo showed the use of a SplitButton on the Application Menu; this demo shows you how to make a much nicer version:

![image](https://github.com/fafalone/UIRibbonDemos/assets/7834493/5359b897-0d70-40ec-b50d-b4ad7329e683)

It also shows making a smaller SplitButton on a regular Tab:

![image](https://github.com/fafalone/UIRibbonDemos/assets/7834493/14d75214-33b9-4946-972a-f2b8150fc275)

Since these are mutually exclusive, when one changes, we need to update the others. So when we get an execute command, we store the selection in a module-level variable, then invalidate the other 3:\
pFramework.InvalidateUICommand IDC_LINESPACE1, UI_INVALIDATIONS_VALUE, vbNullPtr`\
When it's invalidated, it triggers the `RibbonUpdateProperty` event, where we compare them against the module level setting and set their property manually:

```vb6
    If (commandId = IDC_LINESPACE1) Or (commandId = IDC_LINESPACE115) Or (commandId = IDC_LINESPACE115) Or (commandId = IDC_LINESPACE2) Then
        If key Then
            CopyMemory pk, ByVal key, LenB(Of PROPERTYKEY)
        End If
        If IsEqualPKEY(pk, UI_PKEY_BooleanValue) Then
            If commandId = IDC_LINESPACE1 Then
                newValue = IIf(mCurSpacing = LS_1, CVar(True), CVar(False))
            End If
            'repeats for others
```

### CheckBoxes and Column Breaks:
Visible in the picture in the Font Control section a while back, the main tab has a group of CheckBoxes that are simple boolean properties, with a column break and then 3 small icon buttons, using the advanced SizeDefinition fields. You first create a name map, then a custom size definition, which allows Row and ColumnBreak tags; then following that, you just list the commands in order:

```
          <Group CommandName="cmdCheckHdr">
            <SizeDefinition>
              <ControlNameMap>
                <ControlNameDefinition Name="Check1"/>
                <ControlNameDefinition Name="Check2"/>
                <ControlNameDefinition Name="Cut"/>
                <ControlNameDefinition Name="Copy"/>
                <ControlNameDefinition Name="Paste"/>
              </ControlNameMap>
              <GroupSizeDefinition Size="Large">
                <Row>
                  <ControlSizeDefinition ControlName="Check1"/>
                </Row>
                <Row>
                  <ControlSizeDefinition ControlName="Check2"/>
                </Row>
                <ColumnBreak/>
...
            <CheckBox CommandName="cmdCheck1"/>
            <CheckBox CommandName="cmdCheck2"/>
            <Button CommandName="cmdCut"/>
            <Button CommandName="cmdCopy"/>
            <Button CommandName="cmdPaste"/>
```

Omitted: Create the same GroupSizeDefinition for Small and Medium. 

### HelpButton:
The Help Button is now displayed by the minimize ribbon button, up in the top right:

![image](https://github.com/fafalone/UIRibbonDemos/assets/7834493/4200b618-feff-4dcd-908c-4af3de59a535)

This is another command where the image is provided for us. It's a direct child of the `<Ribbon>` group:

```
      <Ribbon.HelpButton>
        <HelpButton CommandName="Help"/>
      </Ribbon.HelpButton>
```


### Context Tabs: 
Astute observers may have noticed an earlier pic with a Tab outlined in green you don't initially see on the form. These are the two 'Context Tabs': Tabs that are only shown when you want to display extra features for some part of your program.  From the Main tab, you can use Select to Show, Unselect to Hide, or Toggle to switch back and forth. These are defined in the `<Ribbon.ContextualTabs>` element. Controlling their visibility is simple, you just set the ContextAvailable property key of their parent command: `pFramework.SetUICommandProperty(IDC_TABTABLE, UI_PKEY_ContextAvailable, vNew)`. This is one place we have to deal with a VT_UI4 Variant. You have to be very careful with these; if VB so much as glances at them it throws an automation error. How it's generally handled in this project, it to manually change the type without touching the data by overwriting the first 2 bytes; tbShellLib has a helper for this replicating an inline available to C/C++ programmers:

```vb6
Public Function InitPropVariantFromUInt32(ByVal ulVal As Long, ppropvar As Variant) As Long
ppropvar = ulVal
Dim vt As Integer = VT_UI4
CopyMemory ppropvar, vt, 2    'Overwrite VT_I4 with VT_UI4
End Function
```
            
### Dropdown Color Pickers:
The Colors tab shows 3 different types of Color Pickers. They each have different preset selections to choose from, and each have a further popup for a full color picker dialog where you can select from the full spectrum.

![image](https://github.com/fafalone/UIRibbonDemos/assets/7834493/235701e5-390f-447f-8c32-13f788b74d80)

The entire popup is generated for you; you need only supply the button and icon for the dropdown.

```
          <Group CommandName="cmdButtonsGroup" SizeDefinition="OneButton">
            <Button CommandName="cmdButtonListColors"/>
          </Group>
          <Group CommandName="cmdDropDownColorPickerGroup" SizeDefinition="ThreeButtons">
            <DropDownColorPicker CommandName="cmdDropDownColorPickerThemeColors"/>
            <DropDownColorPicker CommandName="cmdDropDownColorPickerStandardColors" ColorTemplate="StandardColors"/>
            <DropDownColorPicker CommandName="cmdDropDownColorPickerHighlightColors" ColorTemplate="HighlightColors"/>
          </Group>
```

"Automatic" is an excuse to get the color from Windows with `GetSysColor(COLOR_WINDOWTEXT)`. Otherwise, besides none, an RGB value is specified. It's also a VT_UI4 variant, but it's a COLORREF, so we want the data as-is. Again, we simply overwrite the data type. But this time, we need to get the data from the `commandExecutionProperties`, since the command is the dropdown parent, not an individual button:

```vb6
ElseIf type = UI_SWATCHCOLORTYPE_RGB Then
    Dim vClr As Variant
    If commandExecutionProperties IsNot Nothing Then
        commandExecutionProperties.GetValue UI_PKEY_Color, vClr
        If VariantUI4ToI4(vClr, clr) Then
```

### Icon-only Button Group
![image](https://github.com/fafalone/UIRibbonDemos/assets/7834493/f3aa3f98-9521-436b-8440-4112e06403ec)

Implemented as a Paragraph format set, this shows how to use the advanced SizeDefinition options to make a set of icon-only buttons arranged in rows. One of these  is a dropdown. These buttons are functional and are applied to the selection on the RichEdit control. This is actually one of the harder things to implement in this demo, because icon-only buttons arranged in rows are not the default. This uses a custom size definition similar to the CheckBox / Button group, but defined separately. After the name map, we can specify groups with the IsLabelVisible property set to False:

```
              <ControlGroup>
                <ControlSizeDefinition IsLabelVisible="false" ControlName="ButtonIndent"/>
                <ControlSizeDefinition IsLabelVisible="false" ControlName="ButtonOutdent"/>
              </ControlGroup>
```

We then link it up by name: `<Group CommandName="cmdGroupParagraph" SizeDefinition="ParagraphLayout">`, when listing all the buttons.


### MiniToolbar and Context Popups
![image](https://github.com/fafalone/UIRibbonDemos/assets/7834493/f5791a05-9ad0-4844-831d-3fba0d8ee756)

This feature allows for a popup anywhere on the form of one or both of a mini-toolbar and popup menu. These can have controls like dropdown buttons, toggle buttons, and  more. You can show different combinations of things based on the current view. This  demo has 4 different options you can select from the Colors tab. Click one of the 'Activate Context' buttons, then the CommandButton to show it. The mutually exclusive Toggle Buttons should be easy to understand at this point, they work like the Line Spacing split button, in setting them all after one is clicked. So we'll skip to the popups. After the `</Ribbon>` tag closes out the Ribbon, you can specify `<ContextPopup>` groups, each with a `<ContextPopup.MiniToolbars>` and/or `<ContextPopup.ContextMenus>` group. You then lay out the elements like normal for each, e.g.

```
        <MiniToolbar Name="MiniToolbar3">
          <MenuGroup>
            <Button CommandName="cmdButton1"/>
            <Button CommandName="cmdButton2"/>
            <Button CommandName="cmdButton3"/>
          </MenuGroup>
        </MiniToolbar>

        <ContextMenu Name="ContextMenu1">
          <MenuGroup>
            <Button CommandName="cmdCut"/>
            <Button CommandName="cmdCopy"/>
            <Button CommandName="cmdPaste"/>
          </MenuGroup>
          <MenuGroup>
            <DropDownButton CommandName="cmdMore">
              <Button CommandName="cmdButton1"/>
              <Button CommandName="cmdButton2"/>
              <Button CommandName="cmdButton3"/>
            </DropDownButton>
          </MenuGroup>
        </ContextMenu>
```
Once you done that, you map them to a context to define which one can show. You can have one menu and one minitoolbar per context; you can have one of each but not two or more of any one. 

```
      <ContextPopup.ContextMaps>
        <ContextMap CommandName="cmdContextMap1" ContextMenu="ContextMenu1"/>
        <ContextMap CommandName="cmdContextMap2" ContextMenu="ContextMenu2" MiniToolbar="MiniToolbar2"/>
        <ContextMap CommandName="cmdContextMap3" MiniToolbar="MiniToolbar3"/>
        <ContextMap CommandName="cmdContextMap4" ContextMenu="ContextMenu4"/>
      </ContextPopup.ContextMaps>
```
The id of those commands is what we pass in mCtx to display the popup:

```vb6
        If mCtx = 0 Then Exit Sub
        Dim pt As POINT
        Dim pCtxMenu As IUIContextualUI
        
        If pFramework IsNot Nothing Then
            pFramework.GetView mCtx, IID_IUIContextualUI, pCtxMenu
            If pCtxMenu IsNot Nothing Then
                GetCursorPos pt
                pCtxMenu.ShowAtLocation pt.x, pt.y
            End If
        End If
```
### Recent Items (MRU) List
![image](https://github.com/fafalone/UIRibbonDemos/assets/7834493/e7385015-f8c0-4db3-a2eb-dc5033abd64a)

**NEW** Wanted to add this before the Advanced Demo since the categor header was there. In order to supply the list of items for the Recent Items menu, your UpdateProperties handler gets a request for `UI_PKEY_RelatedItems`. Unfortunately it immediately gets complicated from there. Each item is represented by a class that implements `IUISimplePropertySet`. The old `clsRibbonEvents.twin` has been renamed 'RibbonClasses.twin' and contains `clsRibbonEvents` and a new class: `clsRibbonMRUFile`. This is a generic handler for either files or custom labels. You can specify a file path, and the display name will automatically be looked up for the label, and the type for the label description. Or, you can specify one or both manually. This is all set by taking advantage of twinBASIC's new parameterized constructors:\
`Sub New(sFileFullPath As String, Optional sLabelOverride As String = "", Optional bAutomaticDescriptionOfType As Boolean = True, Optional sLabelDescription As String = "", Optional bPinned As Boolean = False)`\
The way this works is when you use the `New` keyword, you can specify arguments: `Dim pItem As New clsRibbonMRUFile(path, "override", False, "description", False)` rather than have to enter each as a separate property. `IUISimplePropertySet` has only one member: `GetValue`. This part is simple, we just check the PROPERTYKEY it's asking for, and give it the value from the constructor. Only the 3 present are supported; unfortunately you can't provide an icon.

Now that the class is set up, lets look at how to create them in response to the UpdateProperty request and supply them to the Ribbon. The key is looking for a SAFEARRAY of IUnknown, an array of our classes. Since this is a little above tB's native array handling, we'll use the low level `SAFEARRAY` APIs:

```vb6
    Dim pItems() As clsRibbonMRUFile
    ReDim pItems(nMRUItems - 1)
    Dim i As Long
    Dim psa As LongPtr = SafeArrayCreateVector(VT_UNKNOWN, 0, nMRUItems)
    For i = 0 To UBound(pItems)
        Set pItems(i) = New clsRibbonMRUFile("", "Recent file #" & i, , "Description of file " & i, IIf(i = 0, True, False))
        SafeArrayPutElement psa, i, ByVal ObjPtr(pItems(i))
    Next
```
We're not supplying real files for the demo, because I don't know what's where your computer. The first item is pinned just to show that functionality. Now that has to be put into the `PROPVARIANT` that's returned:

```vb6
    Dim ppsa As LongPtr
    VariantSetType newValue, VT_ARRAY Or VT_UNKNOWN
    SafeArrayCopy psa, ppsa
    CopyMemory ByVal PointerAdd(VarPtr(newValue), 8), ppsa, LenB(Of LongPtr)
    SafeArrayDestroy psa
    bSetNewValue = True
```
We manually set it's type, copy the array, and then copy the pointer to the data part of the `PROPVARIANT` before destroying the original.

And that's it... once that's done, we just listen for clicks on `IDC_RECENTITEMS`, our XML command for the list. The item number is given by `VariantUI4ToI4(currentValue, nItem)`. 

## Intro to Galleries Demo

![GalleriesSS](https://i.imgur.com/L81AQSI.gif)

Galleries are one of the nicest features of the Ribbon, but also the most complicated to implement after the Font Control. So before jumping right into the planned Advanced Demo,  which will tread new ground rather than be a simple port of an existing sample, I wanted  to get a handle on basic Gallery use by following the SDK example for them.\
This demo covers 3 types of galleries: In-ribbon, drop down command gallery, and dropdown item gallery. In addition it covers editable and non-editable comboboxes, which are handled the same way as galleries.

These are all a bit of a pain to use as they're mostly populated during runtime, so you need to insert the resources yourself or load them from elsewhere. This can be slow and tedious so for simplicity I just used the existing resource file from the SDK example since this is about how to write the code.

These all work the same basic way: When created, they raise an Update Property request where you're passed an item collect object, then fill it with instances of a class implementing IUISimplePropertySet, which respond to PROPERTYKEYs for images, labels, command ids, command types, and/or category IDs, depending on the specific gallery type. To help with this, a new generic class has been added: clsRibbonGalleryItem. You need only create one of these for each item and use the property lets. For the image, it has helpers to set by either resource id, by HBITMAP, or directly by IUIImage. If you do specify an HBITMAP yourself, note that the Ribbon will take ownership of it; do not free it yourself. 

By this point you should be familiar enough with the XML it's self-explanatory, so let's jump straight to the code.\
For most of these, we're not using categories, and must return `S_FALSE` to indicate that. For Size/Color, we do want categories though. We create a gallery item class using only the labels:

```vb6
        Case IDR_CMD_SIZEANDCOLOR
            If IsEqualPKEY(pk, UI_PKEY_Categories) Then
                Set pCol = currentValue
                
                Dim pSize As New clsRibbonGalleryItem
                pSize.CategoryID = 0
                pSize.Label = LoadStringFromRes(hMod, IDS_SIZE_CATEGORY)
                pCol.Add pSize
                Set pSize = Nothing
                
                Dim pColor As New clsRibbonGalleryItem
                pColor.CategoryID = 1
                pColor.Label = LoadStringFromRes(hMod, IDS_COLOR_CATEGORY)
                pCol.Add pColor
                Set pColor = Nothing
```

A custom API-based function, `LoadStringFromRes`, is used through this project, rather than the intrinsic `LoadResString`, to support running from the IDE or loading from an external DLL like the rest of the Ribbon data. Next we respond to the request for the item source... when we receive this, the `currentValue` has an object implementing `IUICollection`, which we add our items too, again, each an instance of the helper `clsRibbonGalleryItem` helper class:

```vb6
            ElseIf IsEqualPKEY(pk, UI_PKEY_ItemsSource) Then
                Set pCol = currentValue
                Dim scCmdIds(5) As Long
                Dim scCatIds(5) As Long
                
                scCmdIds(0) = IDR_CMD_SMALL
                scCmdIds(1) = IDR_CMD_MEDIUM
                scCmdIds(2) = IDR_CMD_LARGE
                scCmdIds(3) = IDR_CMD_RED
                scCmdIds(4) = IDR_CMD_GREEN
                scCmdIds(5) = IDR_CMD_BLUE
                
                scCatIds(3) = 1
                scCatIds(4) = 1
                scCatIds(5) = 1
                
                For i = 0 To UBound(scCmdIds)
                    Dim pCommand As New clsRibbonGalleryItem
                    pCommand.CategoryID = scCatIds(i)
                    pCommand.CommandID = scCmdIds(i)
                    pCommand.CommandType = UI_COMMANDTYPE_BOOLEAN
                    pCol.Add pCommand
                    Set pCommand = Nothing
                Next
```

All of the galleries, and comboboxes, proceed similarly, each using a slightly different set of properties. For images, the Ribbon Framework uses the `IUIImage` interface and helpfully has a factory to automatically create one from an `HBITMAP`. The helper class further abstracts this away, allowing you to specify just a resource ID. The rest should be pretty straightforward to understand from the code. One oddity encountered... inexplicably, the ComboBoxes were initially way too tiny. This is set by the ribbon, and we use the exact same strings as the C++ version, so it's a complete mystery why our app got tiny ones. Fortunately I was able to find an undocumented workaround; contrary to MSDN documentation, ComboBox controls, not just Spinner controls, receive a request for `UI_PKEY_RepresentativeString`, where you supply a string representing the longest expected contents to be used for sizing it. 


---

That's all for now! Stay tuned for the Advanced Demo, where we'll implement Gallery controls like the Shape Box and Brush dropdown in Paint, ComboBoxes, Spinner Controls, multiple ribbon modes, supplying multiple resolution images for high-DPI and high-contrast alternatives, and more!
