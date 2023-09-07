*** Settings ***
Documentation   Create word documents for compliance and business purposes
...             Read excel file to process each client records
...             Read predefined work template and generate new document with client info
...             Save the document in local folder
...             Send email to the business users with attachments
Library         RPA.Archive
Library         RPA.Excel.Files
Library         RPA.Tables
Library         RPA.Dialogs
Library         RPA.Excel.Application
Library         RPA.FileSystem
Library         RPA.Word.Application
Library         OperatingSystem
Library         String
Library         DateTime

*** Variables ***
${Input_ExcelPath}              ${CURDIR}${/}Input/Company Incorporation Details.xlsx
${StrDefault_CompanyName}       [COMPANY NAME]
${StrDefault_CompanyNumber}     [INSERT REGISTERED COMPANY NUMBER]
${StrDefault_Date}              [Date]
${StrDefault_FullName}          [INSERT DIRECTOR NAME(S)]
${StrDefault_Address}           [INSERT REGISTERED ADDRESS]
${StrDefault_Text}              [is or are]



*** Keywords ***
Read Company info from Excel
        RPA.Excel.Files.Open Workbook    ${Input_ExcelPath}
        ${DT_InputRecords}=     Read Worksheet As Table      name=Summary    header=True    start=4
        Filter empty rows    ${DT_InputRecords}
        Close Workbook
        [Return]    ${DT_InputRecords}

*** Keywords ***
Selected Document
    Add heading    Document to Prepare
    Add drop-down    
    ...  name=document    
    ...  options=Memorandum template   
    ...  default=Memorandum template   
    ...  label=Document
    ${document}     Run dialog
    [Return]   ${document}

*** Keywords ***
Selected Company
    Add heading    Company to Prepare
    Add drop-down    
    ...  name=company     
    ...  options=Fish & Chips LTD.,Baklava A.S.,Olive Oil Essentials S.A.,Sangria S.A.,Limoncello S.p.A.    
    ...  default=Fish & Chips LTD.    
    ...  label=Company
    ${company}    Run dialog
    [Return]     ${company}

*** Keywords ***    
Assigning with row values and replace text
    [Arguments]         ${dtrow}        ${document}        ${company}
    Log     ${dtrow}
    Set Local Variable    ${ID}                     ${dtrow}[ID]
    Log   ${dtrow}[ID]  
    Set Local Variable    ${Company_Name}           ${dtrow}[Company Name]
    Set Local Variable    ${Company_Number}         ${dtrow}[Company Number]
    Set Local Variable    ${Incorporation_Date}     ${dtrow}[Incorporation Date]
    Set Local Variable    ${Address}                ${dtrow}[Registered Address]
    Set Local Variable    ${FullName}               ${dtrow}[Director 1 Full Name] ${dtrow}[Director 2 Full Name]
    Set Local Variable    ${FullName2}              ${dtrow}[Director 2 Full Name]
    ${FullName}=    Replace String    ${FullName}    None    ${EMPTY}
    ${CompanyName}=   Convert To String     ${Company_Name}
    ${Incorporation_Date}=  Convert Date    ${Incorporation_Date}       %d.%m.%Y
    ${Date}=      Convert To String    ${Incorporation_Date}
    ${CompanyNo}=       Convert To String   ${Company_Number}
    IF    "${company}[company]"=="${CompanyName}"
        RPA.Word.Application.Open Application       
        Open File   ${CURDIR}${/}Input/Word_Template/${document}[document].docx
    
        ${documenttocomplete}=      Replace Text      ${StrDefault_CompanyName}        ${Company_Name}
        ${documenttocomplete}=      Replace Text      ${StrDefault_CompanyName}        ${Company_Name}
        ${documenttocomplete}=      Replace Text      ${StrDefault_CompanyNumber}      ${CompanyNo}
        ${documenttocomplete}=      Replace Text      ${StrDefault_Date}               ${Date} 
        ${documenttocomplete}=      Replace Text      ${StrDefault_Address}            ${Address}
        ${documenttocomplete}=      Replace Text      ${StrDefault_FullName}           ${FullName}
        ${documenttocomplete}=      Replace Text      ${StrDefault_Text}               is
        Save Document As    ${CURDIR}${/}Output/${CompanyName}docx
        RPA.Word.Application.Close Document
    END



*** Keywords ***
Get the file to the user
    [Arguments]    ${company}
    Add file    ${CURDIR}${/}Output/${company}[company]docx
    Run dialog


*** Tasks ***
Document Generator
    #Select Document and Company
    ${document}=    Selected Document
    ${company}=    Selected Company
    #Read excel file
    Read Company info from Excel
    Log  Read Excel Completed Successfull
    #Loop through each record and word create template
    ${DT_InputRecords}=     Read Company info from Excel
    FOR    ${row}    IN    @{DT_InputRecords}
        Assigning with row values and replace text       ${row}   ${document}   ${company}
        Log  File creation is done
    END
    Get the file to the user    ${company}