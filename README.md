# weather2excel
This is a program for entering current weather data into an 
Excel spreadsheet (Excel must be installed on your computer)

##Functions
<table> 
  <tr>
    <td><tt>make_worksheet</tt></td>
        <td>Make blank template</td>
          <td>Parameter <tt>-path</tt> | Path of the file to be created (eg. $home\Desktop\test.xlsx)</td>
	     <td>Parameter <tt>-worksheet_name</tt> | Name of the first page of the sheet</td>
	        <td></td>
	           <td></td>
	              <td></td>
  </tr>
  <tr>
    <td><tt>open_and_write</tt></td>
        <td>Enter weather data to Excel file</td>
          <td>Parameter <tt>-path</tt> | Path of the file into which the data is to be entered (eg. $home\Desktop\test.xlsx)</td>
	     <td>Parameter <tt>-row</tt> | Row into which we enter data (e.g. row 1 is the first empty line to enter data)</td>
	        <td>Parameter <tt>-worksheet_name</tt> | Name of the page to which we want to enter data</td>
	           <td>Parameter <tt>-APIKey</tt> | Your API key</td>
	              <td>Parameter <tt>-City</tt> | City you want to enter weather data from</td>
  </tr>
  <tr>
    <td><tt>new_page</tt></td>
        <td>Make new page in worksheet</td>
          <td>Parameter <tt>-path</tt> | Path of the file to which we are adding a new page (eg. $home\Desktop\test.xlsx)</td>
	     <td>Parameter <tt>-worksheet_name</tt> | Name of the new page</td>
	        <td></td>
	           <td></td>
	              <td></td>
  </tr>
 <tr>
    <td><tt>wait_until_hour</tt></td>
        <td>Waits until a given hour or waits until a given hour has passed</td>
          <td>Parameter <tt>-wait_until_hour</tt> | Hour we are waiting for / that is about to pass</td>
	     <td>Parameter <tt>-waiting_until</tt> | If $true - waits for a given hour | If $false - waits for the given hour to pass</td>
	        <td></td>
	           <td></td>
	              <td></td>
  </tr>
  <tr>
    <td><tt>enter_data</tt></td>
        <td>Enters data into a sheet in actual location (if the sheet does not exist, creates one in actual location)</td>
          <td>Parameter <tt>-row</tt> | Row into which we enter data (e.g. row 1 is the first empty line to enter data)</td>
	     <td>Parameter <tt>-worksheet_name</tt> | Name of the page to which we want to enter data</td>
	        <td></td>
	           <td></td>
	              <td></td>
  </tr>
</table>

##How to use?
Paste the code from functions.ps1 into the powershell console, press Enter. 
Now you can enter the above commands into the powershell console.

##Example usage
Create a folder on your desktop called "Weather", then copy the code from code.ps1 and then paste it into the powershell console.
The program is designed to operate 24 hours a day, 7 days a week. At the following hours (5, 9, 13, 17, 21) he enters the appropriate data into the sheets.
