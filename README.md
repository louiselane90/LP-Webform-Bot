# LP-Webform-Bot
Excel macro to transfer data from inventory sheet to entries into List Perfectly

## Project Description
Data entered into inventory tracking worksheet in Excel is auto entered into inventory via web form at www.listperfectly.com

Uses the following applications:
Selenium, Excel VBA Macros

## Development
When macro is ran, data from "Resale Inventory" is input into a www.listperfectly.com catalogue listing. 									
The Identifier in column A requires a '1' to tell the macro to upload the information in this row. 									
Once it is uploaded, the macro marks it as 'Completed.' 									
The macro launches Google Chrome and requires the Selenium reference to be enabled in VBA.									
Selenium GoogleChrome version must stay compatible with your Google Chrome version.									
https://github.com/florentbr/seleniumbasic/releases									
									
To use:									
Download Selenium									
Edit macro and enter your username and password for List Perfectly									
