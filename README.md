# Subaward Budget Reviewer

Using Visual Studio or VS Code, write a small .NET console application (using .NET 6 or .NET 7) that: 

·        Reads all spreadsheets from a folder (use the 3 attached Excel spreadsheets). 

·        For each file, output to the console the file name followed by each subrecipient name from that file. The subrecipient names will be under “G. Other Direct Costs” in the format “Subaward: {SubRecipientName}” 

·        Finally, output a distinct list of all subrecipients along with the total subaward amount that subrecipient received across all files. 

  

Requirements for the above .NET application: 

·        The app should work with any spreadsheet in this format, and be able to support a variable number of subaward rows (0 – unknown). 

·        The app should have at least one unit test defined that confirms there are 4 subrecipients “Indiana”, “Mayo”, “Purdue”, and “Florida” in SubawardBudgetExample1.xlsx. 

·        The app should be checked into a publicly accessible Github or Azure Devops repository that the reviewers can pull and run (both the console app and unit test), without any modification.  
