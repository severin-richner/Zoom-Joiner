# Zoom-Joiner

(only works on Windows)

This script automatically joins Zoom meetings given a weekday and a time. No account needed. 
The script saves the meetings and assumes these are repeated weekly. Joining is done using the Zoom meeting link and uses the default browser.


## Instructions:

1)	install python (version 3.8 or newer) and the Zoom application

	[Python](https://www.python.org/downloads/) (make sure to select "Add Python X to PATH"), [Zoom](https://zoom.us/support/download)
	

2)	install python packages via Command Line

	> pip install pywin32 keyboard winshell

3)	make sure your standard Browser has pop-ups enabled (for opening Zoom via link)

	(should work fine with: Chrome, Brave, Edge)

4)	download the files "icon.ico" and "Zoom-Joiner.py" and put together in some folder of your choice

Done! Just open/run the "Zoom-Joiner.py" file with python.

#### Sidenote: If you don't know the link for the Zoom meeting, the usual format is:

_https://zoom.us/j/{MeetingID}_ or _https://zoom.us/j/{MeetingID}?pwd={Password}_

Created by Severin Richner
