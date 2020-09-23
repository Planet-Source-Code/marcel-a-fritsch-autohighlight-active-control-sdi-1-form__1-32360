<div align="center">

## Autohighlight active control \(SDI/1 Form\)


</div>

### Description

This is a very simple and useful solution to highlight input controls without writting a function for each control. I use the SetWindowsHookEx function to receive all Windows messages in a callback function (in this example WindowProc). The WindowProc function calls the SETKILLFocus function if there is an WM_SETFOCUS or WM_KILLFOCUS message and delivers the

the Handle of the control to the SETKILLFocus function.In this Example the highlighting is done only for textboxes and comboboxes but you can easily change the SETKILLFocus function to process other types of controls.

I think the explanation in the source code should answer all other questions. Please vote, if you think its a good solution.
 
### More Info
 
When running this progam in the IDE do not use the STOP-Button to exit the program, because the unhook function will not be executed and the IDE crashes!!!


<span>             |<span>
---                |---
**Submitted On**   |2002-03-05 13:47:20
**By**             |[Marcel A\. Fritsch](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marcel-a-fritsch.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Autohighli59466352002\.zip](https://github.com/Planet-Source-Code/marcel-a-fritsch-autohighlight-active-control-sdi-1-form__1-32360/archive/master.zip)








