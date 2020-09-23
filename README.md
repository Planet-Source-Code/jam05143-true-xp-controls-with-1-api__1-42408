<div align="center">

## True XP Controls, with 1 API


</div>

### Description

Objective:

To create and use a manifest file for displaying common controls WinXP style

on a machine running WinXP with the XP theme enabled.

Function:

This code will only provide WinXP controls on a machine running

WinXP. The XP theme must be enabled (Not the classic theme) for

this code to work properly. Also, this only works when the program

is compiled as an executable, and will not work in the IDE environment.

Once the InitXP sub is called, all forms and controls (the main controls

and any controls in the "Microsoft Windows Common Controls 5.0 (SP2)" ocx)

on all forms will be display in winXP style, so you only have to call the

sub once. Also, MsgBox's and InputBox's will have the xP theme. What's cool

about this code is if the xP theme is enabled, but a different color scheme is

used (other than the default blue), the manifest automatically acounts for that.

The truely great thing about this is that no new activex controls are introduced,

and therefor you do not need to account for different OS. If you want XP controls

on machines other than XP, there are hundreds of controls on the net. This code

reduces deployment costs, is small, and is self-contained.

I have not tested the code on any machine other than winxp, but this should

not generate any errors. If it does, a simple call to "GetVersionEx" could

show which OS is running, and in the Sub InitXP, add the code:

If WinXP = True Then

all current code

End If

Known Problems:

Command Buttons, Option Buttons, and Frames do not display properly when

placed inside of another Frame. Currently working on this problem.

One last note:

The exe will not display XP controls the first time it

is run, and I have not found a way around that yet . . . Maybe creating

a program that first creates the manifest, then shell's the main program

before ending . . . Or maybe the install program could solve it somehow

Credit:

I got the idea from this code from another source, I believe on planet

source code, but the code did not work, and was not well organized/commented.

I have no idea who the genious was that came up with this code, but all the

credit in the world goes to that person.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-01-14 01:19:54
**By**             |[jam05143](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jam05143.md)
**Level**          |Beginner
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[True\_XP\_Co1528861142003\.zip](https://github.com/Planet-Source-Code/jam05143-true-xp-controls-with-1-api__1-42408/archive/master.zip)








