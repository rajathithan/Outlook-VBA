# Extract Employee details for analysis using Outlook VBA

### The extracted module can be used for only testing purposes, You can run the code at your own risk. (I am not liable for it :) ! )

This module uses AddressEntries object model and the GetFirst & GetLast methods of Outlook to extract the emloyee details like - 
company name, first name,last name,department,title,office ,city, alias, email address, supervisor's -(firstname,lastname,
emailaddress,alias).

You can import this module to your outlook vba editor

The extracted information is stored in an exccel sheet under C drive. 

I have made the settings in such a way that, it will extract only 10 entries, as the program is very slow and can extract only 
50 to 60 contact information per minute. 



