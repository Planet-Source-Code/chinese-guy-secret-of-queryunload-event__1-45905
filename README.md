<div align="center">

## Secret of QueryUnload event


</div>

### Description

Determine when your app is being closed by Windows OS, or by Task Manager or by user.

This can be useful when you have to do cleaning up before your app exits.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chinese Guy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chinese-guy.md)
**Level**          |Beginner
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chinese-guy-secret-of-queryunload-event__1-45905/archive/master.zip)





### Source Code

<p>'///////////////////////////////////////</p>
<p>'Excerpt from MSDN documentation.</p>
<p>'///////////////////////////////////////</p>
<p>'vbFormControlMenu 0 : The user chose the Close command from the Control menu on the form.
</p>
<p>'vbFormCode 1 : The Unload statement is invoked from code.</p>
<p>'vbAppWindows 2 : The current Microsoft Windows operating environment session is ending.
</p>
<p>'vbAppTaskManager 3 : The Microsoft Windows Task Manager is closing the application.</p>
<p>'Remarks</p>
<p>'This event is typically used to make sure there are no unfinished tasks in the forms
'included in an application before that application closes.
'For example, if a user has not yet saved some new data in any form,
'your application can prompt the user to save the data.</p>
<p>'When an application closes, you can use either the QueryUnload or Unload
'event procedure to set the Cancel property to True, stopping the closing process.
'However, the QueryUnload event occurs in all forms before any are unloaded,
'and the Unload event occurs as each form is unloaded.</p>
<p>Private Sub Command1_Click()</p>
<p>   Unload Me</p>
<p>End Sub</p>
<p>Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)</p>
<p>   Select Case UnloadMode</p>
<p>   Case vbFormControlMenu</p>
<p>      'The user clicked the Close button on the upper-right of form</p>
<p>      If MsgBox("UnloadMode : Close button. Wanna exit?", vbYesNo) = vbNo Then
         Cancel = True</p>
<p>      End If</p>
<p>   Case vbFormCode</p>
<p>      'There's Unload statement in the code.</p>
<p>      If MsgBox("UnloadMode : Unload statement. Wanna exit?", vbYesNo) = vbNo Then
         Cancel = True</p>
<p>      End If</p>
<p>   Case vbAppWindows</p>
<p>      'Windows OS session is ending.</p>
<p>      If MsgBox("UnloadMode : Windows OS. Wanna exit?", vbYesNo) = vbNo Then
         Cancel = True</p>
<p>      End If</p>
<p>   Case vbAppTaskManager</p>
<p>      'Windows Task Manager is closing this app.</p>
<p>      If MsgBox("UnloadMode : Task Manager. Wanna exit?", vbYesNo) = vbNo Then
         Cancel = True</p>
<p>      End If</p>
<p>   End Select</p>
<p>End Sub</p>

