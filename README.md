# windowsninja

This is a windows tool that can be starter and run on a Windows Client/Server to monitor windows running on the Console.
I wrote this tool to monitor a large number of Windows clients running a Navision client that did not have any log handling.

The tool is written in "Autoit". 

Filestructure and explanation:

clientexe  : Folder containing executables and Source code. Start the application using "WindowsNinja.bat"
Config     : Folder containing WindowsNinjaConfig.csv with parameters for the application. The filter "VISIBLE_TEXT" can be simple or REGEX. Edit in Excel (!).
Logs       : Will contain logs for the client running the application
pictures   : Will contain pictures from the client running the application. Also snapshot of Errors
README.md  : This Readme file
LICENSE.MD : License



Quickstart:
1. Copy the structure to a folder locally or if several computers to a fileshare
2. Review settings in Config\WindowsNinjaConfig.csv, but I recommend to leave as-is for simple testing.
3. Run the clientexe\WindowsNinja.bat on the client. 
4. Test by starting Notepad and write "Testerror"
5. Default setup will pause WindowsNinja as long as there is mouse activity on the console. Rightclick on the icon on the taskbar and select "Pause/Off(Startnow)
6. You should observe activity under the folders "Logs" and "pictures"
7. Review settings in the Config\WindowsNinjaConfig.csv for you're needs. Specially note the timer polling intervals "StatusCaptureTimer" to avoid too much polling activity.

RDP/Console
The code uses commands to windows "qwinsta" to throw any RDP session to the Console (or snapshot will not work). This should be OK if you only have one desktop in use on the server/client, but note that the console will be active. This will not work if you have several RDP sessions you want to monitor (one console per machine)- Protect the computer from physical access to the Console! 


Known limitations
1. More documentation 
2. The code is not perfect, but it should work. I am not a professional coder(!)
3. WindowsNinjaConfig.csv is not an optimal configuration file structure.
4. Routines for cleaning up slack messages could have been implemented. Currently all messages posted will be kept. Its good for a historic view though.
5. Error history for snapshots might have been implemented. 
6. Currently only present errors are shown, once solved it is deleted. This is fine to get a "right now" glimpse of errors in the environment though.
7. For some windows, you might get two or more messages for the same window.

Further documentation:
I will try to create some more documentation for it eventually, but feel free to contact me on tonnyrh@gmail.com if you have any questions.


Disclaimer:
In no event shall the author be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use this software or documentation, even if the author has been advised of the possibility of such damages.



Tonny Roger Holm
2019
