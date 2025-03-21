**Summary**

This script will use Graph API to Search against SharePoint Online and OneDrive to locate content using search. Since we are using Application based permissions, you don’t need explicit access to the site to retrieve results.

The Azure APP will need the following permisions:


![image](https://github.com/user-attachments/assets/7624bdb6-62b4-4b9c-ad02-0ca58cd0fc8d)

The actual search query is embedded into the **$requestPayload** parameter:


![image](https://github.com/user-attachments/assets/32748983-e0b1-493f-9b59-67ee7206839e)


The  "$searchQuery" will look at a file in c:\temp\userlist.txt by default to locate the actual query to issue.

Example of query for a OneDrive site:

"(Contoso Purchasing Data - Q11.xlsx) (path:\"https://m365x49978400-my.sharepoint.com/personal/admin_m365x49978400_onmicrosoft_com\")"

![image](https://github.com/user-attachments/assets/afc3eeaf-56ed-4afb-93e0-e31feb989bbb)


**!Important!**

The searchRegion must match your Tenant regions.

_Performing a search with application permission is limited to one geographic region. You must specify a value for the region property in your search request when you use application permissions to run a search. To get the region value for a tenant, use the dataLocationCode property of the siteCollection resource._


After running the script, the output will create a CSV file in the user local temp folder %temp%.

![image](https://github.com/user-attachments/assets/7a383b28-3beb-459f-b3d2-ee822b819501)


Example Output:

![image](https://github.com/user-attachments/assets/e071d9ef-cd77-4476-9921-1d9f2d571ee2)


**More Information:**

Use the Microsoft Search API to query data

https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview?view=graph-rest-1.0

searchEntity: query

https://learn.microsoft.com/en-us/graph/api/search-query?view=graph-rest-1.0&tabs=http

Search SharePoint content with application permissions

https://learn.microsoft.com/en-us/graph/search-concept-searchall
