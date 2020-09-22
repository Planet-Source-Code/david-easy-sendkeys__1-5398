<div align="center">

## Easy SendKeys


</div>

### Description

How to use ALT & Function Keys with the sendkey command
 
### More Info
 
1 Command Button

Copy the code into the Command1_Click Event.

Calculator is installed on the PC


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david.md)
**Level**          |Beginner
**User Rating**    |3.2 (16 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[DDE](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/dde__1-28.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-easy-sendkeys__1-5398/archive/master.zip)





### Source Code

```
Dim ReturnValue, I
ReturnValue = Shell("CALC.EXE", 1)	' Run Calculator.
AppActivate ReturnValue 	' Activate the Calculator.
For I = 1 To 100	' Set up counting loop.
SendKeys I & "{+}", True	' Send keystrokes to Calculator
Next I	' to add each value of I.
SendKeys "=", True	' Get grand total.
SendKeys "%{F4}", True	' Send ALT+F4 to close Calculator.
```

