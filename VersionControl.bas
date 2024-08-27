Attribute VB_Name = "VersionControl"
Option Explicit
'@Change this before Production

Private Const VersionNumber = "24.05.16"
Private Const PreviousVersion = "23.11.22"
Public Const TestStatus As Boolean = False

Private Function getVersionNumber()
    getVersionNumber = VersionNumber
End Function

Public Sub displayVersion()
    MsgBox "Payroll version: " & Words.vLine(2) & getVersionNumber(), vbInformation, "Version Control"
End Sub

'### PREVIOUS VERSIONS ### 'The error found and cause are for the version before it, the resolution is placed in the following version
'VER 25.05.16
'1. Replaced shift planning token field with humanity password field and made loginToM.ShiftPlanning use the Humanity password
'VER 23.11.22
'1. Extended every account sheet to allow for 100 employees.
'VER 23.05.01 ->
'1. Added rows to 'Total'. Max is now 115 employees.
'VER: 22.07.15 ->
'1. Fixed missing OT accounts and ordered them according to worksheet order
'2. Removed Driver, Raley Field and Sutter
'3. Removed Lead shifts from OT as well as moved PC into normal category
'   a. There was once a need to separate these because of separate OT pay but this was never meant to be the case. Separate pay is tips and OT is at Regular rate
'VER: 22.05.16 ->
'1. Added more columns to HC sheet
'   a. Notes: You can just insert the columns and the name manager will update the ranges automatically.
'VER: 21.04.26 ->
'1. No shift lead bonus calculated
'    a.
'    -For some reason IHTMLElement.FirstChild() is not the same as IHTMLElement.Children(0) anymore.
'    -It's also possible Humanity removed a tr element changing the position of the targeted element. Either way...
'    Solution: Changed all instances of FirstChild to Children(0)
'    Solution (Dev only): For testing purposes, don't attempt shift lead calculation for future dates from today (there were no events 4-1 to 4-15 so I pushed payroll to 4-16 to 4-30)
'        This causes an error if there is no one scheduled yet, and it's easier to account for dev testing than assume management would ever try to pull time sheets early.
'    -Some arrays that were supposed to reference parking control referenced private party (this workbook might be a previous version as I thought I fixed this years ago).
'Solution:     Corrected References
    
'    b.
'    -Getting a warning about shift leads not being found despite the bonus being calculated
'    -This was due to Humanity changing classNames to having more values than just the date. So you could not find lead shifts based on the class of the valet shift.
'Solution:     Don 't use the class of the valet shift. Use the basic classSelector of just the date. This will target both.
    
'2. New account "Scott's Seafood Roundhouse" not importing
'    -This is due to Humanity not listing the whole restaurant name if it's too long
'    -ex: Scott's Seafood Roundhouse -> Scott's Seafood Roun..
'    Solution: if it's too long (and *only* if it's too long), the html will have an attribute "title" or "oldtitle" with the full name. Use that instead.

'3. Won't export to ADP
'    -ADP changed their login form
'    Solution: Update code to accomodate ADP login form inputs and logic
'    -ADP is not receiving payroll inputs. End result is that payroll spreadsheet total does not match ADP total.
'    -Name matching was comparing first middle last to last, first.
'    Solution: Swap last and first so they are compared correctly
    
'4. Does not work with Macs
'    -I do not have a Mac to test any issues and iOS is incompatible with Internet Explorer
'    Solution: It might be possible to convert scripting to use Chrome, but the effort is impossible as I don't have a Mac
'    Note: ADP warns that it will eventually be incompatible with IE. Converting to Chrome is inevitable.
    
'5. Need to hit "apply" button when importing from Humanity on the time sheets page. It won't happen automatically
'    -This was due to Humanity changing their formatting from "Apr 15, 2021" to "04/15/2021"
'    Solution: Change the URL being generated programatically
    
'6 (found by dev). Formatting on 'Total' sheet dragged blue down for some reason
'    -No idea why the conditional formatting changed. End user should not be able to do this
'    Solution: Fixed conditional formatting.
'7 OT was calculating the validation by averaging everyone's wages and multiplying it by the number of OT hours
'    -This calculation was susceptible to error given a variance in people's wages.
'    -It was also obsolete because OT Hours are entered into ADP, not as CC Tips as it was before.
'    Solution: Added a row to OT that just calculates hours.

'VER: 18.07.03 ->
'   UPDATE: Added Burgers & Brewhouse, Arden Fair mall, Season's 52, and The Cheesecake Factory
'   UPDATE: Alphabetized Accounts
'VER: 18.05.17 ->
'   1)
'   ERROR: Mismatch Error when exporting to ADP
'   CAUSE: Row on 'Total' missing OT formula
'   RESOLUTION: Added formula
'VER: 18.04.27 ->
'   Made OT calculate correct hours based off all OT hours pay 1.5 * hourly rate
'   Made Regular calculate correct hours based off subtracting OT hours
'   Fixed Dates on OT sheet to be "N/A" when they day of the month is not in the pay period. Hopefully Steve never puts N/A in PP or PC
'   Removed T22
'VER: 18.04.04 ->
'   1)
'   ERROR: Shift Lead bonses not getting pulled
'   CAUSE: Humanity changed the way they structure the data. className by date is reformatted and text is mashed together
'   RESOLUTION: Reformat how it's formed.
'   IMPORTANT: Shift Lead count calculates correctly but only gives minimum bonus because counting valets not resolved
'   2)
'   UPDATE: Change Esquire to The Diplomat, Pink Martini to Lucca, Add Mikuni
'   RESOLUTION: Accounts added.
'VER: 18.02.02 ->
'   ERROR: Total sheet was pulling 2nd PP EG hours for PM
'   CAUSE: =SUMIF(PM!$1:$1,$AI4,IF(DAY(PayDay1)=1,EG!$17:$17,EG!$34:$34)) <- "EG" in second half of formula
'   RESOLUTION: Replace all "EG" in PM column with "PM
'VER: 18.01.26 ->
'   ERROR: Hours not carrying over to Total
'   CAUSE: Column A names were not being formatted right
'   RESOLUTION: Reformat names in a stored variable so that they can be entered on Total in correct format and then compared using the Import Format which is stored in the variable.
'VER: 18.01.18 ->
'   UPDATE: Steve requests breaks aren't deducted since they're usually paid breaks and he can deduct lunches himself somehow
'   RESOLUTION: deduct_breaks flag changed from 1 to 0 in Humanity URL extension
'VER: 17.12.18 ->
'   1)
'   ERROR: Invalid procedure call fixing name formatting in ShiftLeadFromSP
'   CAUSE: Payroll always assumes Last, First format, sometimes employees are First Last
'   RESOLUTION: Account for times when name is formatted incorrectly
'   2)
'   ERROR: Browser stall when logging in to Shift Planning before getting time sheets
'   CAUSE: Would not navigate to right URL after log in
'   RESOLUTION: If having to log in, wait 10 seconds and then re-navigate to correct URL
'VER: 17.11.20 ->
'   ERROR: ADP crashes on payroll info fill page
'   CAUSE: They changed the locator of the page number
'   RESOLUTION: Update the locator and method of parsing the page number
'VER: 17.09.25 ->
'   UPDATE: House of Oliver and TopGolf are being renamed to Canon East Sacramento and Tres Hermanas
'   RESOLUTION: Create a beta user form that allows USA Valet to make these changes themselves
'VER: 17.09.01 ->
'   ERROR: Some employees being skipped on export to ADP
'   CAUSE: Not enough sleep time (200 ms)
'   RESOLUTION: additional sleep (500 ms)
'VER: 17.07.17 ->
'   1)
'   ERROR: SP not given enough time to load
'   CAUSE: Not enough sleep time
'   RESOLUTION: additional sleep
'   2)
'   ERROR: OT and UV not enough columns
'   CAUSE:
'   RESOLUTION: Add more columns
'   3)
'   ERROR: OT not calculating all hours
'   CAUSE: Bad formula
'   RESOLUTION:
'   ADDED FEATURE: User added to Import sheet time tracker.
'VER: 17.06.19 ->
'   1)
'   ERROR: Not going to correct URL for humanity timesheets
'   CAUSE:
'   RESOLUTION:
'   2)
'   ERROR: Not correctly copying over employee names from Import spreadhsheet to accounts spreadsheets
'   CAUSE:
'   RESOLUTION:
'VER: 17.05.18 -> Changed humanity login button locator to //button[@name='login'][1] to //button[@name='login' and text()='Log in']
'   1)
'   ERROR: Putting CC Tips in Salary column on ADP
'   CAUSE: sendKeys(tab) causing mouse to find wrong cell in ADP
'   RESOLUTION: sleep(100) after each tab to give the browser time to catch up to the new click.
'   2)
'   ERROR: Login button in Humanity login page not getting clicked
'   CAUSE: Humanity changed the actual login button to be the second index
'   RESOLUTION: Verify the text of the button is "Log in" as the first one has no text.
'VER: 17.05.02 ->

'VER: 17.04.05 (Not released)
