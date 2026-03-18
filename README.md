# DMI Install Instructions
## Step #1 - Download Files
Download the release files from [SelectionInfo/SelectionInfo/bin/Release](https://minhaskamal.github.io/DownGit/#/home?url=https://github.com/rainbow-of-gravity/Inventor_Substation_Addin/tree/master/SelectionInfo/SelectionInfo/bin/Release)

## Step #2 - Place Files in Correct Location
.dll files goes here ->     C:\Program Files\Autodesk\Inventor 2022\Bin  
NOTE: There are multiple .dll files, also if .stdole is already in this folder it is recommended to NOT REPLACE

.addin file goes here  ->   C:\Users\\[USER]\AppData\Roaming\Autodesk\Inventor 2022\Addins
Update your the username in the above path though ^^^^^^^^^^^^
 
## Step #3 - Start
 Start Inventor

## Step #4 - Set Keybinds
To set the keybinds go to Tools > Customize > Then set the two commands "000 Go Down Hierarchy" and "000 Go Up Hierarchy" to whatever keybinds you want. I suggest "Alt + A" and "Alt + Q" respectively

# HOW TO USE
In a open assembly
Add a new window pane by clicking the "+" at the top of the Document Tree Viewer and selection "Selection2"
This creates a new window pane that can be moved around. When clicking on different parts ensure that cursor is in "Select Part Priority" mode (google how to do this).
To go up and down use the buttons in the window ribbon - if the buttons are not there, keybinds can still be used. Google how to change keybinds in inventor
















# IGNORE BELOW!!!!!!!
## Autodesk-Inventor (Vestigial Instructions - copied over from cloned repository)

<b>Tools for Autodesk Inventor</b>

<b>SelectionInfo</b> is a sample add-in application for Autodesk Inventor by CAD Studio; it displays an iProperties palette as a modeless, dockable UI element in Inventor. This palette displays document iProperties (e.g. Part Number, Mass, Area, etc) of any user selected component. The user can select the component in either the window or in the browser tree. You can change, add, or remove which iProperties are shown in the pallete by editing the <i>SelectionInfoSelector.cs</i> file.

<img src="SelectionInfo/SelectionIP.gif">


<b>Installation</b>
 - Refer to <a href="SelectionInfo/SelectionInfo/install.txt">install.txt</a> for installation instructions.
- A precompiled DLL of the base add-in is included (Inventor 2018 or higher) <a href="Autodesk-Inventor/SelectionInfo/SelectionInfo/bin/Release">here</a>.
 
<b>Customization</b>
 - Refer to <a href="SelectionInfo/Readme.md">Readme.md</a> for customization instructions and an explanation of the code structure.


Contact CAD Studio at <a href="https://www.cadstudio.cz">www.cadstudio.cz</a> or <a href="https://www.cadforum.cz">www.cadforum.cz</a> or Facebook: <a href="https://www.facebook.com/CADstudio">@CADstudio</a> for more information.
