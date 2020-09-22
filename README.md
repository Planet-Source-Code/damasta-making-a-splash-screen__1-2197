<div align="center">

## Making A Splash Screen\!\!


</div>

### Description

Load any program and the first thing you see is a splash screen. Splash screens usually display the program's name along with a graphic of some sort, such as a pen for the Microsoft Word splash screen or the company's logo for the Procomm Plus splash screen.
 
### More Info
 
At the start of every project you create, use Form1 (the first form to appear when you run your program) as your splash screen. Place the Timer control on the form and define the Interval property of the Timer. (An interval of 1000 equals 1 second and an interval of 65,535 is approximately 1 minute. You may have to experiment with the proper interval to display for your splash screen.)

Then set the following Form1 properties to False: ControlBox, MaxButton, and MinButton, since you don't want users to be able to minimize or maximize your splash screen.

Next, draw an Image box (Image1) so it completely covers Form1 and paste in the graphic image you want to display on your splash screen. For simple splash screen images, you can use the Paintbrush program that comes with Windows or any of the numerous commercial or shareware paint programs available.

Finally, add the following code to the Timer1 event procedure:


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[DaMaStA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/damasta.md)
**Level**          |Unknown
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/damasta-making-a-splash-screen__1-2197/archive/master.zip)





### Source Code

```
Sub Timer1_Timer ()
 Unload Form1
 Load Form2
 Form2.Show
End Sub
'The above code tells Visual Basic that after the Timer control waits for the time specified by the Interval property, it should unload Form1 (your splash screen) and then load and display Form2 (which contains your 'program's first screen).
Sub Image1_Click ()
 Timer1.Enabled = False
 Unload Form1
 Load Form2
 Form2.Show
End Sub
'Experiment with creating splash screens for all of your programs. Whether you need to use splash screens to give your programs a professional look, or you need a splash screen to fake the illusion of speed for your users, you may find splash screens to be a simple, yet effective way to make your programs more acceptable to the average user.
The above steps are all that you need for a splash screen to disguise the fact that your program takes time to load. However, if your program loads quickly, you can give the user the option of clicking on the splash screen to make it go away, rather than wait a specified amount of time.
To give your splash screens the ability to disappear when the user clicks the mouse, add the following code to the Image1 event procedure stored on Form1
```

