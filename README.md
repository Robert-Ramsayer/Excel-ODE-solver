# Excel-ODE-solver
Runga-Kutta ODE function installed into excel

Welcome, to start with, this prgram has a GNU license attached to it. Please read the document for the exact details. For a breif summery, the program is free to download and use but you need to give me a citation. 

==========

Purpose
This program isn't to just add Runge-Kutta ODE solver to excel, I've seen that and it's ussually a clunky mess. This is adding a macro into excel that makes Solving ODE's in excel easy and fast. Within excel there's a program named Visual Basic, you can access it by opening up an excel sheet and pressing ALT+F11. A new page will open up and you can program new functions into excel. I programmed an differential equation solver to impress a proffessor. 

=============
Advantages
This program is undoubtedly less accurate than an ODE solver in something like Python, but it's a lot faster. In my use, I found no delay between inputing new values and getting results. I suggest you use it for hunting for values as an estimator, and either tolerating some error or using a more advanced program to fine tune the results.

===============

How to download

In order to use this just download the excel file on my page, and open it up. When downloading it you may get a notice that there may be a virus in the file. There isn't, I wouldn't put my name on it if there was. If your still concerned, there's another option I'll bring up in the next paragraph. When you open the excel file Runge-Kutta.xlsm, you'll get a notice at the top of the  excel sheet stating 'Security Warning. Macros have been disabled. Options' My program is a Macro, and this program is not dangerous, other might be. Click on the Options button. You'll see 2 options: 'Help protect me from unknown content (recomended)' and 'Enable this content.'  To use my program you'll need to enable it. If you click the other option, you can redo it at any time just by closing the excel sheet and opening it back up again.  As for why Excel is so concerned, they can't predict what will be in every program you opening up and they'd rather just not chance it. 

Now if you're still suspecious, there is another option. I copied down the VBA program in the Runge-Kutta Program document. You can just open up an Excel sheet, press ALT+F11, right click on one the the excel sheets, insert, new macro, and then copy and paste my code. You'll miss some examples but that's it. 

==============

How to use
There are 2 differential equation solvers: 1st and 2nd order Runge-Kutta. Let's start with the 1st order. To use this, first you'll need to open up the VBA program. Again, open up program by opening up the Excel sheet with the macro enabled and then hitting ALT+F11, and go to the function at the top starting titled f(). Enter the equation in the form of y'= the rest of the equation, then go back to the excel sheet. In my example spread sheet you'll see that there is a row of 4 letters with 4 numbers under it: h, t, y, y'. Do the same, and under that row, put your starting values; the h is your step, t is initial time, y is your starting y value, and y' doens't matter just put a number. In the row under that select an area 3 columns wide and 1 row deep. Type '=rkf(' input the values in this order ' h, and in one highlighted suqare t, y, y' ( it should looke like this: =rkf($B$3,C3:E3). Keep the $, they're imporitant. Then hit CTRL+SHIFT+Enter; this special command is for array outputs. The output will be the next iteration of the ODE. You can continue to more iterations by highlighting this new row and dragging it down by the square in it's bottom left corner. 

To use the 2nd order one, the process is very similar. Go into the program and change the G() instead. Inter the equation in the form if y''= the rest of the euation. z=y'. Save it, go to the excel sheet and enter the variable in a row in this order: h,t,y,y',y''. Enter those values, y'' just has to be a number. Highlight the row under in a 4x1 arangement, then enter '=rks(' meaning Runge-Kutta Second order. Enter the new values by clicking on h's value, and then highlighting the values from t-y''. Finish the () and use enter the same way as before. Should look like this 'rks($B$24,C26:F26).' You've now got the next iteration. Keep going to make it however long you need. 

=========== 

Modifying to your needs

I strongly suspect that if you're looking for this program you can solve a simple ODE faster than it would take you to download this program, and read my document. Don't worry, this program can be adjusted easily to meet your needs. I suspect your trying to hunt down the right coefficient for an ODE to do soemthing. If that is the case, go to the RKS function. See how h is intered, and just add another variable. Then scroll down to see how t is entered into G(), and just add your variable. If that's what your doing, you'll find your value pretty quick. Also use CTRL+ENTER; you can enter a variable without switching cells. 
