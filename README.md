# <h1>Hours Registration VBA APP</h1>
Excel VBA APP to register hours on a weekly basis with secutiry and password protection. The App enables saving the data to separate a database file. 

<img width="600" alt="image" src="https://user-images.githubusercontent.com/19918869/165584944-dee22d6c-0a16-40a3-b854-6540c17a1c02.png">

<h2>Functionalities of the app:</h2>
 <div> 1 - Password protected so that people cannot manipulate prior filled in hours </div>
 <div> 2 - Functionality to add more people to the App </div> 
 <div> 3 - Time limit so that App will close after a certain amount of time </div> 
 <div> 4 - Saving of a weekly hours PDF file </div>
 <div> 5 - Saving of hours to seperate database </div> 

<h2>The App Contains: </h2>
<div>1 urenReg.xlsb : this is the landing page where people can access the controls of the app (such as add users and register hours). Furthermore, main controls are also located here. </div> 
<div>2 masterData.xlsb : seperate database file where on a day-by-day basis hours are stored  </div> 
<br></br> 

<h2>Script urenReg.xlsb:</h2>
<h3>Workbook modules</h3>
<li> Workbook_Open() : lock specific columnswith a password protection before open based on the current day of the week </li>
<li> Workbook_BeforeClose(): lock all the columns </li>
<li> Workbook_SheetChange(): call upon time action that closes the workbook after 10 minutes </li>

<h3>Modules</h3> 
-dataTrans(): requiring password and performs file transfer to database, saves a pdf of week, and cleans the current week 
- clearfillinData(): function to clear data 
- speed(): function to speed up code execution by disabeling certain visual features 
- slow(): function to go back to default settings 
- TimeSetting(): function to close and save workbook after specified amount of time
- 

<h2>Script masterData.xlsb:</h2>
