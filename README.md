# Excel-ODE-solver
Runga-Kutta ODE function installed into excel

Welcome, to start with, this prgram has a GNU license attached to it. Please read the document for the exact details. For a breif summery, the program is free to download and use but you need to give me a citation. 

2nd, the purpose of this program isn't to just add Runge-Kutta ODE solver to excel, I've seen that and it's ussually a clunky mess. This is adding a macro into excel that makes Solving ODE's in excel easy and fast. Within excel there's a program named Visual Basic, you can access it by opening up an excel sheet and pressing ALT+F11. A new page will open up and you can program new functions into excel. I programmed an differential equation solver to impress a proffessor. 

In order to use this just download the excel file on my page, and open it up. When downloading it you may get a notice that there may be a virus in the file. There isn't, I wouldn't put my name on it if there was. If your still concerned, there's another option I'll bring up in the next paragraph. When you open the excel file Runge-Kutta.xlsm, you'll get a notice at the top of the  excel sheet stating 'Security Warning. Macros have been disabled. Options' My program is a Macro, and this program is not dangerous, other might be. Click on the Options button. You'll see 2 options: 'Help protect me from unknown content (recomended)' and 'Enable this content.'  To use my program you'll need to enable it. If you click the other option, you can redo it at any time just by closing the excel sheet and opening it back up again.  As for why Excel is so concerned, they can't predict what will be in every program you opening up and they'd rather just not chance it. 

Now if you're still suspecious, there is another option. I copied down the VBA program in the Runge-Kutta Program document. You can just open up an Excel sheet, press ALT+F11, right click on one the the excel sheets, insert, new macro, and then copy and paste my code. 

==============

How to use
There are 2 differential equation solvers: First and Second order Runge-Kutta. Let's start with the first order Runge-Kutta Solver. To use this, first you'll need to open up the VBA program. Again, open up program by opening up the Excel sheet with the macro, and go to the function at the top starting titled f(). Enter the equation in the form of y'= the rest of the equation, then go back to the excel sheet. In my example spread sheet you'll see that there is a row of 5 letters with 4 numbers under it: h, t, y, y'. Under that row, put your starting values; the h is your step, put your initial t value under t, your y(t) value under y, and y' just has to be a number. In the row under that select an area 3 columns wide and 1 column deep. Type '=rkf(' input the values in this order ' h, (and in one highlighted suqare ) t, y, y' ( it shoudl looke like this 
