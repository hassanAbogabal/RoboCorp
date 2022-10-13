*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.
Library    RPA.Browser.Selenium  auto_close=${False}
Library    RPA.HTTP
Library    RPA.Excel.Files
Library    RPA.RobotLogListener
Library    RPA.PDF
Library    RPA.Desktop
Library    RPA.Dialogs



*** Tasks ***
 Insert the sales data for the week and export it as PDF
   Open the intranet website
   Log in 
   Download the Excel file
   Fill the form using the data from the Excel File 
   Collect the results
   Export the table as a PDF
   [Teardown]     Log out and close the browser



*** Keywords ***
Open the intranet website
   Open Available Browser    https://robotsparebinindustries.com/


Log in 
  Input Text    username    maria
  Input Password    password    thoushallnotpass
  Submit Form
  Wait Until Page Contains Element    id:sales-form

Download the Excel file 
  Add heading    Please Enter URL of Oreder CSV File.. 
  Add Date Input    Url    
  ${URL} =    Run dialog
  Download     ${URL}    overwrite=True


Fill and submit the form for one person 
  [Arguments]    ${Sales_reps}
  Input Text    firstname    ${Sales_reps}[First Name]
  Input Text    lastname    ${Sales_reps}[Last Name]
  Input Text    salesresult    ${Sales_reps}[Sales]
  Select From List By Value   salestarget   ${Sales_reps}[Sales Target]
  Click Button    Submit

   
Fill the form using the data from the Excel File 
  Open Workbook    SalesData.xlsx
  ${Sales_reps}=   Read Worksheet As Table    header=True
  Close Workbook
  FOR    ${sales_rep}    IN    @{Sales_reps}
     Fill and submit the form for one person    ${sales_rep}
      
  END

Collect the results
  Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Export the table as a PDF
  Wait Until Element Is Visible    id:sales-results
  ${sales_result_html}=    Get Element Attribute    id:sales-results    outerHTML
  Html To Pdf    ${sales_result_html}    ${OUTPUT_DIR}${/}sales_results.pdf
  
  

  
Log out and close the browser
    Click Button    Log out
    Close Browser



