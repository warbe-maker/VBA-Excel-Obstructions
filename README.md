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
Download and import [_mObstructions_][1] or have a look at the complete [development and test Workbook][2] which also provides an unattended self asserting progression test. All can be found in the corresponding public [GitHub Repository][2]

## Usage
### The _All_ service
For a 'brute force' approach _mObstructions.All_ performs all the other services allowing to skip one or another. _All_ is an 'all but ...' service because any obstruction may be ignored.  

The _All_ service is used as follows
```vb
mObstructions.All obs_service:=enEliminate, obs_ws:=<worksheet-object>, obs_range:=<range-object>
    ' any code which otherwise cannot be executed 
mObstructions.All obs_service:=enRestore, obs_ws:=<worksheet-object>, obs_range:=<range-object>
```

The _All_ service has the following named arguments:

|    Part              | Description                    |
| -------------------- |------------------------------- |
| obs_service          | Enumerated expression enEliminate or enRestore |
| obs_ws               | Expression identifying a Worksheet object |
| obs_range            | Range expression pointing to the current selected cells of those cells which should be considered for an un-merge |
| obs_.....            | Boolean expression, optional, default to True. When set to False the corresponding obstruction is ignored. |

All the other public obstruction services may also be performed individually as an alternative to the 'all but ...' approach.


[1]:https://gitcdn.link/repo/warbe-maker/Common-Excel-VBA-Obstructions-Services/master/source/mObstructions.bas
[2]:https://gitcdn.link/repo/warbe-maker/Common-Excel-VBA-Obstructions-Services/Obstructions.xlsm
[3]:https://github.com/warbe-maker/Common-Excel-VBA-Obstructions-Services