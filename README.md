# Excel VBA Obstructions Services
Provides the means to temporarily _Eliminate_ and finally _Restore_ hurdles like _Protected Sheets_, _Filtered Rows_, _Hidden Columns_, _Merged Cells_, and _Application Events_ within any procedure.

## The approach
The _Eliminate_ service pushes the information of an obstructions on a Worksheet specific stack from where the _Restore_ service pops it. This ensures the restoration of exactly that status which had been eliminated - provided these two services are performed strictly paired. The stacking approach allows _Eliminate_/_Restore_ services being performed in any number  at any nested level without worrying about the status **provided the two services are performed strictly paired**.

## Obstructions serviced
### Application Events
One of the most common obstructions potentially causing loops or unintended operations when content of cells is changed.

### Sheet Protection
Protected sheets are serviced provided they are not protected by a password.

### Filtered Rows and Hidden Columns
The _Eliminate_ service saves the current status into a CustomView by stacking a temporary name and sets AutoSave off as well as displays hidden columns. The _Restore_ service just re-displays the CustomView identified by the temporary name popped from the stack.
I've found ample discussions in forums how to save AutoSave properties with enormous complex solutions all finally stating that there might be something missing. For me there's nothing better than using a system function to solve that problem.

### Merged Areas
When it comes to merged cells experienced VBA developers say: "Don't use them. They create endless headache.". I am very much focused on sheet designs which not only do meet user requirements but also provide best practice layout. I felt very sorry about the idea that merged cells must not be used in order to avoid subsequent VBA coding problems.
The _Eliminate_ service un-merges all merged areas which do relate to a given range. I.e. merged cells which relate to given ranges rows or columns. The _Eliminate_ service copies the content of the merged area into each un-merged cell. So no matter which column or row may be removed the range can be re-merged by the _Restore_ service. No hassle ever again.

## Installation
1. Download [_mObstructions.bas_][1] or have a look at the complete [development and test Workbook][2] which also provides an unattended self asserting progression test.
2. Import the _mObstruction.bas_ 
3. In the VBE add a Reference to the "Microsoft Scripting Runtime"

Have a look at the corresponding public [GitHub Repository][2]

## Usage
### The _All_ service
For a 'brute force' approach _All_ service performs all the other individual obstruction services in an all but ... manner. I.e. individual obstructions by default included may be exempted.  

The _All_ service is used as follows
```vb
mObstructions.All obs_service:=enEliminate, obs_ws:=<worksheet-object>, obs_range:=<range-object>
    ' any code which otherwise cannot be executed 
mObstructions.All obs_service:=enRestore, obs_ws:=<worksheet-object>, obs_range:=<range-object>
```

The _All_ service has the following named arguments:

|    Part              | Description                    |
| -------------------- |------------------------------- |
| _obs\_service_          | Enumerated expression, either _enEliminate_ or _enRestore_ |
| _obs\_ws_               | Expression identifying a Worksheet object. The Worksheet for which the obstructions are to be eliminated or retored. |
| _obs\_range_            | Range expression, obtional, defaults to Nothing. The range is only obligatory when the MergedAreas obstruction is performed. The range usually points to the current selected cells of which the rows and columns are considered relevant for any merged area's un-merge. |
| _obs\_\<obstruction>_     | Boolean expression, optional, default to True. When set to False the corresponding obstruction is ignored. |

### The individual obstruction services
Any individual obstruction service may be performed in any procedure provided eliminate and restore is performed strictly paired.  

| Service Name             | Service |
|--------------------------|---------|
| _ApplEvents_             | Eliminate (turn off) and Restore (turn back on) Application.EnableEvents |
| _MergedAreas_            | Eliminate (un-merge) and Restore (re-merge) merged areas concerned by a specific selected range. |
| _SheetProtection_        | Eliminate (un-protect) and Restore (protect) an individual Worksheet |
| _FilteredRowsHiddenCols_ | Eliminate (turn off _AutoFilter_, display hidden columns) and Restore (re-show a temporary created CustomView) of an individual Worksheet. It should be noticed that the means CustomView is not Worksheet specific but has a Workbook scope. Though AutoSave and hidden columns are managed for a specific Worksheet only the CustomView concerns all sheets. This is perfectly implemented however. |

### The _Rewind_ service
This service typically performed after an error condition to ensure that no obstruction is left un-restored. However, this service must only be used in the entry procedure, i.e. the procedure which performed an obstruction eliminate  in the first place. Under regular conditions, when code execution ends in the entry procedure this service will have nothing to be performed and thus will not do any harm. A typical usage scenario in the entry procedure may look as follows:

```VB
    On Error Goto eh
    mObstructions.All enEliminate, ThisWorksheet
    ' in here may be any number of nested sub procedures
    ' which also perform individual obstruction services e.g. ' to ensure sheet protection of a specific sheet is turned ' off - and at the end of it restored 
    mObstructions.All enRestore, ThisWorksheet
    
xt: mObstructions.Rewind
    Exit Sub

eh: 'Display error message
    Goto xt ' perform a clean exit
End Sub
```

[1]:https://gitcdn.link/cdn/warbe-maker/Common-Excel-VBA-Obstructions-Services/master/source/mObstructions.bas
[2]:https://gitcdn.link/cdn/warbe-maker/Common-Excel-VBA-Obstructions-Services/Obstructions.xlsm
[3]:https://github.com/warbe-maker/Common-Excel-VBA-Obstructions-Services