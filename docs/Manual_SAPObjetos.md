



# SAP
  
With this module you can automate SAP R3 and Hana. Perform every action that you can record with the tracker created by Stefan Schnell.  

*Read this in other languages: [English](Manual_SAPObjetos.md), [Português](Manual_SAPObjetos.pr.md), [Español](Manual_SAPObjetos.es.md)*
  
![banner](imgs/Banner_SAPObjetos.png)
## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  

## How to use this module
In order to use this module, you have to connect to your SAP Account.

You have to activate the scripts of SAP GUI.

In SAP GUI, you should go to RZ11 transaction; in the name param, write "sapgui/user_scripting" and click in "Display" (as shown in images).

![sap1](imgs/sap1.png)
![sap2](imgs/sap2.png)

Click on "Change Value and on "New Value" select "TRUE". Save it.

![sap3](imgs/sap3.png)
![sap4](imgs/sap4.png)

The next image shows that the GUI scripting is activated in RZ11.

![sap5](imgs/sap5.png)

To activate scripting in SAP GUI you have to go to "Personalize" and go to "Options".

In "Accessibility & Scripting, go to "Scripting" and activate "Enable scripting" option. Save it.

![sap6](imgs/sap6.png)

How to work properly?

Open Tracker.exe

![sap7](imgs/sap7.png)

It will open a window like this

![sap8](imgs/sap8.png)

Open SAP and log in. Once logged, click on this icon and will sync with the active session on SAP.

![sap9](imgs/sap9.png)

Go to 
"Recorder" and click on the Python icon.

![sap10](imgs/sap10.png)


## Description of the commands

### LoginSap
  
Open SAP application, establish the connection and login
|Parameters|Description|example|
| --- | --- | --- |
|Path SAP|SAP Application path|C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe|
|Connection Name|Connection name where you want to do login|Developer (Direct)|
|User ID|The ID of SAP User field |wnd[0]/usr/txtRSYST-BNAME|
|User|User name that you enter in the user field to login|User1|
|Password ID|The ID of SAP password field |wnd[0]/usr/pwdRSYST-BCODE|
|Password|Password that you enter in the password field to login|Password|
|Timeout|Time lapse (in seconds) to search for the element|10|
|Save connection result|Variable where the connection result will be saved|variable|

### Connect
  
Connect to an opened SAP application
|Parameters|Description|example|
| --- | --- | --- |
|Connection Name|Connection name where you want to do login|Developer (Direct)|
|Save connection result|Variable where the connection result will be saved|variable|

### Execute an action
  
Execute an action in SAP such as selecting, focusing or modifying the text of an element through different properties (setFocus, text, etc.)
|Parameters|Description|example|
| --- | --- | --- |
|Object ID|ID obtained in the tracker|wnd[0]/tbar[0]/okcd|
|Row number (Optional)|Number of the row where you want to execute the action. In the tracker it will appear as selectedRows|0|
|Column name (Optional)|Name of the column  where you want to execute the action. In the tracker it will appear as currentCellColumn|TYPE_DOC|
|Value|Value shown in the tracker after an equal (=) (eg id.attribute=value), or as in parentheses (eg id.attribute(value)|QMNUM|
|Option|Propery indicated in the tracker after findById('some id'). Ex session.findById('wnd[0]/tbar[0]/okcd').property|Option|
|Timeout|Time lapse (in seconds) to search for the element|10|

### Extract Text
  
Extract the text of an object in SAP, using the text property
|Parameters|Description|example|
| --- | --- | --- |
|Object ID|ID obtained in the tracker|wnd[0]/tbar[0]/okcd|
|Timeout|Time lapse (in seconds) to search for the element|10|
|Assign to Variable|Variable name where the result will be saved|variable|

### Check/Uncheck
  
Mark or unmark an object in SAP
|Parameters|Description|example|
| --- | --- | --- |
|Option|Select check or uncheck|Uncheck|
|Check individually with absoluteRow|Option to check them with absolute row|Checkbox|
|Check all|Option to check them all|Checkbox|
|Absolute Row|Column to interact with|3|
|Object ID|ID obtained in the tracker|wnd[0]/tbar[0]/okcd|
|Timeout|Time lapse (in seconds) to search for the element|10|

### Extract Cell
  
Extract the text in a cell, it can be with GetItemText or GetCellValue
|Parameters|Description|example|
| --- | --- | --- |
|Object ID|ID obtained in the tracker|wnd[0]/tbar[0]/okcd|
|Row number|Number of the row where you want to extract the text. In the tracker it will appear as selectedRows|0|
|Column name|Name of the column where you want to extract the text. In the tracker it will appear as currentCellColumn|TYPE_DOC|
|Option||Option|
|Timeout|Time lapse (in seconds) to search for the element|10|
|Assign to Variable|Variable name where the result will be saved|variable|

### Click Cell
  
Click on a cell. It can be with clickCurrentCell, setCurrentCell or doubleClickCurrentCell
|Parameters|Description|example|
| --- | --- | --- |
|Object ID|ID obtained in the tracker|wnd[0]/tbar[0]/okcd|
|Row number|Number of the row where you want to extract the text. In the tracker it will appear as selectedRows|0|
|Column name|Name of the column where you want to extract the text. In the tracker it will appear as currentCellColumn|TYPE_DOC|
|Click type|Propery indicated in the tracker after findById('some id'). Ex session.findById('wnd[0]/tbar[0]/okcd').property|Option|
|Timeout|Time lapse (in seconds) to search for the element|10|

### Run script
  
Execute VBS script recorded with SAP
|Parameters|Description|example|
| --- | --- | --- |
|Path VBA Script|SBS file path|C:/path/to/script.vbs|
|Default Script||Option|

### Get state checkbox
  
Return True if checkbox is selected
|Parameters|Description|example|
| --- | --- | --- |
|Object ID|ID obtained in the tracker|wnd[0]/tbar[0]/okcd|
|Timeout|Time lapse (in seconds) to search for the element|10|
|Assign to Variable|Variable name where the result will be saved|variable|

### Send key
  
Replicate send key event
|Parameters|Description|example|
| --- | --- | --- |
|Object ID|ID obtained in the tracker|wnd[0]/tbar[0]/okcd|
|Row number (Optional)|Number of the row where you want to send the key. In the tracker it will appear as selectedRows|0|
|Column name (Optional)|Name of the column where you want to send the key. In the tracker it will appear as currentCellColumn|TYPE_DOC|
|Key|Key or combination to send|F1|
|Timeout|Time lapse (in seconds) to search for the element|10|

### Get Property
  
Get a property of the SAP Object especified
|Parameters|Description|example|
| --- | --- | --- |
|Object ID|ID obtained in the tracker|wnd[0]/tbar[0]/okcd|
|Property|Property to get. To see all property, go to https//help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/760.03/en-US/a2e9357389334dc89eecc1fb13999ee3.html|Option|
|Timeout|Time lapse (in seconds) to search for the element|10|
|Assign to Variable|Variable name where the result will be saved|variable|

### Wait for object
  
Wait for object to be visible
|Parameters|Description|example|
| --- | --- | --- |
|Object ID|ID obtained in the tracker|wnd[0]/tbar[0]/okcd|
|Timeout|Time lapse (in seconds) to search for the element|10|
|Assign to Variable|Variable name where the result will be saved|variable|
