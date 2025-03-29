This program is made to input the messages sent to patients from Practice Mojo into Softdent as those programs don't allow intercompatibility.
This is made for Oak Meadow Dental Center but can be used by everyone who has a practice mojo account as well as those who use softdent. 

Code length has been reduced greatly since the start, but there are still ways to improve efficiency of the program.

Username and password are not hardcoded and use a config.json file to store it instead. For example:
{
    "USERNAME": "username",
    "PASSWORD": "password"
}

For ahk = AHK(), it only works if you already have auto hotkey installed on your computer. If you don't, you have to pip install ahk_binary and use ahk = AHK(executable_path = "PATH").

To run: open the file through python and type in the date in the format requested (MM/DD/YYYY) and press the start button. Make sure softdent is open into the patient search window and make sure chrome is closed.

You can pause and resume by clicking scroll lock, and you can force exit the program with num lock.

Requried packages:
pip install ahk_binary
pip install selenium
pip install pywinauto
pip install keyboard
pip install tflite
