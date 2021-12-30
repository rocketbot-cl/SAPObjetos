



# SAP
  
With this module you can automate SAP R3 and Hana. Perform all the actions that you can record with the tracker shared 
by Stefan Schnell in the official SAP forum.

## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  


## How to use this module
In order to use this module, you have to connect to your SAP Account.

You have to activate the scripts of SAP GUI.

In SAP GUI, you should go to RZ11 transaction; in the name param, write "sapgui/user_scripting" and click in "Display" (as shown in images).

![sap1](docs/imgs/sap1.png)
![sap2](docs/imgs/sap2.png)


Click on "Change Value and on "New Value" Select "TRUE". Save it.

![sap3](docs/imgs/sap3.png)
![sap4](docs/imgs/sap4.png)

The next image shows that the GUI scripting is activated.

![sap5](docs/imgs/sap5.png)

To activate scripting in SAP GUI you have to go to "Personalize" and go to "Options".

In "Accessibility & Scripting, go to "Scripting" and activate "Enable scripting" option. Save it.

![sap6](docs/imgs/sap6.png)

How to work properly

Open Tracker.exe

![sap7](docs/imgs/sap7.png)

It will open a window like this

![sap8](docs/imgs/sap8.png)

Open SAP and log in. Once logged, click on this icon and will sync with the active session on SAP.

![sap9](docs/imgs/sap9.png)

Go to "Recorder" and click on the Python icon.

![sap10](docs/imgs/sap10.png)


## Overview


1. LoginSap  
Open the sap application, establish the connection and login

2. Click on object  
Click on SAP. It can be through different properties, such as setFocus, text, etc. You can also modify the text of an 
element

3. Extract Text  
Extract the text of an object in sap, using the text property

4. Check/Uncheck  
Mark or unmark an object in SAP

5. Extract Cell  
Extract the text in a cell, it can be with GetItemText or GetCellValue

6. Click Cell  
Click on a cell. It can be with clickCurrentCell, setCurrentCell or doubleClickCurrentCell

7. Run script  
Execute VBS script recorded with SAP

8. Get state checkbox  
Return True if checkbox is selected

9. Send key  
Replicate send key event

10. Get Property  
Get a property of the SAP Object especified  




----
### OS

- windows
- docker

### Dependencies

### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)