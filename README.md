# weather2excel
This is a program for entering current weather data into an 
Excel spreadsheet (Excel must be installed on your computer)

## Functions
<table> 
  <tr>
    <td><tt>make_worksheet</tt></td>
        <td>Make blank template</td>
          <td><tt>-path</tt> | Path of the file to be created (eg. $home\Desktop\test.xlsx)</td>
	         <td><tt>-worksheet_name</tt> | Name of the first page of the sheet</td>
	            <td></td>
	               <td></td>
	                  <td></td>
  </tr>
  <tr>
    <td><tt>open_and_write</tt></td>
        <td>Enter weather data to Excel file</td>
          <td><tt>-path</tt> | Path of the file into which the data is to be entered (eg. $home\Desktop\test.xlsx)</td>
	         <td><tt>-row</tt> | Row into which we enter data (e.g. row 1 is the first empty line to enter data)</td>
	            <td><tt>-worksheet_name</tt> | Name of the page to which we want to enter data</td>
	               <td><tt>-APIKey</tt> | Your API key</td>
	                  <td><tt>-City</tt> | City you want to enter weather data from</td>
  </tr>
  <tr>
    <td><tt>new_page</tt></td>
        <td>Make new page in worksheet</td>
          <td><tt>-path</tt> | Path of the file to which we are adding a new page (eg. $home\Desktop\test.xlsx)</td>
	         <td><tt>-worksheet_name</tt> | Name of the new page</td>
	            <td></td>
	               <td></td>
	                  <td></td>
  </tr>
 <tr>
    <td><tt>wait_until_hour</tt></td>
        <td>Waits until a given hour or waits until a given hour has passed</td>
          <td><tt>-wait_until_hour</tt> | Hour we are waiting for / that is about to pass</td>
	         <td><tt>-waiting_until</tt> | If $true - waits for a given hour | If $false - waits for the given hour to pass</td>
	            <td></td>
	               <td></td>
	                  <td></td>
  </tr>
  <tr>
    <td><tt>enter_data</tt></td>
        <td>Enters data into a sheet in actual location (if the sheet does not exist, creates one in actual location)</td>
          <td><tt>-row</tt> | Row into which we enter data (e.g. row 1 is the first empty line to enter data)</td>
	         <td><tt>-worksheet_name</tt> | Name of the page to which we want to enter data</td>
	            <td></td>
	               <td></td>
	                  <td></td>
  </tr>
</table>

## Change data language
Change line 89 (```$URL = "https://api.weatherapi.com/v1/current.json?key=$APIKey&q=$City&aqi=yes"```) to this:
```$URL = "https://api.weatherapi.com/v1/current.json?key=$APIKey&q=$City&aqi=yes&lang=YOUR_LANGUAGE"```, 
where **YOUR_LANGUAGE** is your langcode

**All available langcodes are in the table below:**

<table>
    <tr>
        <th>Language</th>
        <th>lang code</th>
    </tr>
    <tr>
        <td>Arabic</td>
        <td>ar</td>
    </tr>
    <tr>
        <td>Bengali</td>
        <td>bn</td>
    </tr>
    <tr>
        <td>Bulgarian</td>
        <td>bg</td>
    </tr>
    <tr>
        <td>Chinese Simplified</td>
        <td>zh</td>
    </tr>
    <tr>
        <td>Chinese Traditional</td>
        <td>zh_tw</td>
    </tr>
    <tr>
        <td>Czech</td>
        <td>cs</td>
    </tr>
    <tr>
        <td>Danish</td>
        <td>da</td>
    </tr>
    <tr>
        <td>Dutch</td>
        <td>nl</td>
    </tr>
    <tr>
        <td>Finnish</td>
        <td>fi</td>
    </tr>
    <tr>
        <td>French</td>
        <td>fr</td>
    </tr>
    <tr>
        <td>German</td>
        <td>de</td>
    </tr>
    <tr>
        <td>Greek</td>
        <td>el</td>
    </tr>
    <tr>
        <td>Hindi</td>
        <td>hi</td>
    </tr>
    <tr>
        <td>Hungarian</td>
        <td>hu</td>
    </tr>
    <tr>
        <td>Italian</td>
        <td>it</td>
    </tr>
    <tr>
        <td>Japanese</td>
        <td>ja</td>
    </tr>
    <tr>
        <td>Javanese</td>
        <td>jv</td>
    </tr>
    <tr>
        <td>Korean</td>
        <td>ko</td>
    </tr>
    <tr>
        <td>Mandarin</td>
        <td>zh_cmn</td>
    </tr>
    <tr>
        <td>Marathi</td>
        <td>mr</td>
    </tr>
    <tr>
        <td>Polish</td>
        <td>pl</td>
    </tr>
    <tr>
        <td>Portuguese</td>
        <td>pt</td>
    </tr>
    <tr>
        <td>Punjabi</td>
        <td>pa</td>
    </tr>
    <tr>
        <td>Romanian</td>
        <td>ro</td>
    </tr>
    <tr>
        <td>Russian</td>
        <td>ru</td>
    </tr>
    <tr>
        <td>Serbian</td>
        <td>sr</td>
    </tr>
    <tr>
        <td>Sinhalese</td>
        <td>si</td>
    </tr>
    <tr>
        <td>Slovak</td>
        <td>sk</td>
    </tr>
    <tr>
        <td>Spanish</td>
        <td>es</td>
    </tr>
    <tr>
        <td>Swedish</td>
        <td>sv</td>
    </tr>
    <tr>
        <td>Tamil</td>
        <td>ta</td>
    </tr>
    <tr>
        <td>Telugu</td>
        <td>te</td>
    </tr>
    <tr>
        <td>Turkish</td>
        <td>tr</td>
    </tr>
    <tr>
        <td>Ukrainian</td>
        <td>uk</td>
    </tr>
    <tr>
        <td>Urdu</td>
        <td>ur</td>
    </tr>
    <tr>
        <td>Vietnamese</td>
        <td>vi</td>
    </tr>
    <tr>
        <td>Wu (Shanghainese)</td>
        <td>zh_wuu</td>
    </tr>
    <tr>
        <td>Xiang</td>
        <td>zh_hsn</td>
    </tr>
    <tr>
        <td>Yue (Cantonese)</td>
        <td>zh_yue</td>
    </tr>
    <tr>
        <td>Zulu</td>
        <td>zu</td>
    </tr>
</table>

## How to use?
Paste the code from functions.ps1 into the powershell console, press Enter. 
Now you can enter the above commands into the powershell console.

## Example usage
Copy the code from code.ps1 and then paste it into the powershell console. That's all you need to do :)
**The program is designed to operate 24 hours a day, 7 days a week.** At the following hours (5, 9, 13, 17, 21) he enters the appropriate data into the sheets.
