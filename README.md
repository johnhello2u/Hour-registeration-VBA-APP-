# <h1>Hours Registration VBA APP </h1>
Excel VBA APP to register hours on a weekly basis with secutiry and password protection. The App enables saving the data to a separate database file, adding new names, saving pdfs, and going though different weeks. 

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
-dataTrans(): General Function that requires a password and performs file transfer to database, saves a pdf of week, and cleans the current week.
- clearfillinData(): Function to clear data 
- speed(): Function to speed up code execution by disabeling certain visual features 
- slow(): Function to go back to default settings 
- TimeSetting(): Function to close and save workbook after specified amount of time
- SelectSheetsToPrint(): Function saves the current current sheet in pdf format with the corresponding week as name
- savewb(): Function used to perform a save of workbook and saving the inputted name to an audit trail. 
- createOutputSheet(): Function creates a sheet with the data prepped in a table for transfer to the external database
- deleteDataInput(): Function that deletes the transfer sheet 
- add_name(): Function that let users add new names to the hour registration form 
- Button6_Click(): Function to go forward 1 week with the dates 
- Button7_Click(): Function to go back 1 week with the dates 

<h2>Script masterData.xlsb:</h2>
