________________________________________________________________________________________________

»VB Screen Saver« --> By Alex M
________________________________________________________________________________________________


________________________________________________________________________________________________

Creating the Screen Saver
________________________________________________________________________________________________

Its pretty easy, just follow these steps,
	1. Compile the Code to an .Exe file
	2. Rename the .Exe to a .Scr
	3. Right click the .Scr file and select install
	4. Done.
	5. If you wan't you could copy the .Scr file to C:\windows\system32



________________________________________________________________________________________________

Before hand
________________________________________________________________________________________________

	This was originally a little muck around project with the mouse. I made the original program so that wherever the mouse moved it drew little circles or dots, etc. But then I remembered that upon opening a previous .scr file with a hex viewer, that it started with the characters "MZ". Now this is when you say, "So What?". Well This is important because normal Executables also start with "MZ" and thats when it hit me, What if I rename the .exe to .scr? Well to some of you more advanced programmers/computer users may be laughing at this point saying "Yeah, I already knew that you could rename the file from .exe to .scr" but It was a learning experience and I'm not even a real programmer as such, programming is just an interest not my profession. 

	Anyways, Back to something more relevant, I also used the command$ string to find out that windows sends data to the program to tell it whether it wants to change settings "/c" or to test settings "/p" or to run the actual screen saver "Null". Another problem was the fact that whenever the program started, it started minimised so I loaded a separate form that was maximised. Sorry about the poor structure of the code, I learned basic using GWBasic as many of its elements of its structure still exist in my code (such as GOTO, String$, Integer%). I Hope this code is some use to you, you CAN use any part of code you want, Copy what you want, delete what you don't.


________________________________________________________________________________________________

Notes/Explanations
________________________________________________________________________________________________

	There are several Zip files included. The Zip Files ending with the word 'Backup' are the main screen saver program with remarks in the code. The Zip Files ending with the word 'Extra' are a few quick Screen saver program I made, They use the same layout as The main screen saver but use different algorithms for the Screen_Saver_Form form. The 3rd option in the Star Field Screen saver is pretty good, and the Matrix screen saver is pretty cool too (Check them out!!).

	Some programmers don't know what the %, # and $ signs are. Well to most programmers you should know what they stand for! But to my supprise, I was showing some code to a awesome programmer, who uses such things as Directx api calls in his programs, and He was at a loss as to what these signs stand for. Well it is simple - Integer%, Number(includes decimals)#, LongInt&, String$. The use of these signs also REDUCES the SIZE of your COMPILED programs!! The use of 'Str$' instead of 'Str' and 'Chr$' instead of 'Chr', and so on, Also reduces Compiled program file size!! So, Change the 'Dim String$' to 'Dim String as String' if you must, but you are making your programs unnecessarily bigger (Try it for yourself!!).

	The Other 'Extra' Zip files have no remarks, Thats why they are labeled 'Extra'. They are in the same layout as the Main Screen saver, but use different algoritms for the Screen_Saver_Form form as mentioned above.

	Anyway, I Hope this code is of some use to you, and I also hope you enjoy using this Code. If not, You have several Cool screen savers you can use!! Have Fun!!


________________________________________________________________________________________________

Creating the Screen Saver
________________________________________________________________________________________________

Its pretty easy, just follow these steps,
	1. Compile the Code to an .Exe file
	2. Rename the .Exe to a .Scr
	3. Right click the .Scr file and select install
	4. Done.
	5. If you wan't you could copy the .Scr file to C:\windows\system32


________________________________________________________________________________________________

Trouble Shooting
________________________________________________________________________________________________
	Note, when compiling, if you encounter errors, change the compiling method to P-code instead of native code (Native code is faster, but can cause problems)

	Any Further Questions, E-mail me at Alex_murray1@hotmail.com, I'll do my best to answer them (Also Make the subject relevant or I'll think its spam, and I check my e-mail atleast every week)


________________________________________________________________________________________________

Disclaimer
________________________________________________________________________________________________
This program has No Warranty of any sort, Any use of the Software is at your own risk.
Do remove this text file