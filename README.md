# A robot that prepares your documents for you ;)

## Powered by [Capto](https://www.wearecapto.com)

> This robot needs to be run in a Windows environment.

<img src="Input\images\DocmentGenerator.jpg" style="margin-bottom:20px;width:800px;">

This robot completes the documents you need to complete in seconds. Just need to tell him the template and the company to prepare it for and it does the rest :)

The workflow involves working with a Windows application, a data source, and a series of dialogue windows.

This robot is written in Python. It uses Python standard libraries and some [RPA Framework](https://rpaframework.org/) libraries.

Here is the main "task" definition:

```rpaframework
*** Tasks ***
Document Generator
    #Clear directories
    Enter authorization code
    RPA.FileSystem.Empty Directory     Output
    RPA.FileSystem.Empty Directory     Archive
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
    Archive Word into ZIP files
```

To use it on your own, you will have to make the following adjustments:

- Gather your client list and required information for completing the documents (Name, Address, Shareholders, etc).
- Select and adapt your word document with dynamic fields for completion ("Name Field", "Address Field", "Shareholder 1 Field", etc).
- Adapt the Fields within the robot framework to select the information of your client and write them in the selected word.

The Windows application executable is included in the repository for convenience.

> See the full code for implementation details!

Capto team is constantly looking for improvement, so please, don't hesitate to leave your review on the following [link](https://forms.gle/vrJCXqMZj6dyHnyt5).
