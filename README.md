# Description
Allows users to connect to a SharePoint Site using the API Client Secret and:
- Obtain lists of files recursively from a SharePoint site directory
- Download / copy files from SharePoint site (preserving or flattening hierarchy)
<br>

## Set-up
### SharePoint Access
In order to access SharePoint via the API, an app-principle will need to be setup.

This can be done by following the guide [here](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs).
>__Summary__
>
>Navigate to a site in your tenant (e.g. https://contoso.sharepoint.com) and then call the appregnew.aspx page (e.g. https://contoso.sharepoint.com/_layouts/15/appregnew.aspx). In this page click on the Generate button to generate a client id and client secret and fill the remaining information like shown in the screen-shot below. [Link](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs#setting-up-an-app-only-principal-with-tenant-permissions)

<br>

The guide by Martin Noah also provides information about how to set-up the SharePoint API Access - [here](https://martinnoah.com/sharepoint-rest-api-with-python.html)

<br>

## SharePointClient
Notes:
- The SharePoint in Microsoft 365 application principal ID is always 00000003-0000-0ff1-ce00-000000000000. This generic value identifies SharePoint in Microsoft 365 objects in a Microsoft 365 organization. [Link](https://learn.microsoft.com/en-us/sharepoint/hybrid/configure-server-to-server-authentication#:~:text=The%20SharePoint%20in%20Microsoft%20365%20application%20principal%20ID%20is%20always%2000000003%2D0000%2D0ff1%2Dce00%2D000000000000.%20This%20generic%20value%20identifies%20SharePoint%20in%20Microsoft%20365%20objects%20in%20a%20Microsoft%20365%20organization.)
- The Client ID and Client Secret (obtained via set-up) will be used when initialising the class.
<br>

## Additional Notes
You will need to setup the SharePoint API access for _each_ site that you want to get information/data from. And obtain the Client ID/Secret for each site.
>For example:
>
>If your SharePoint site URL is 'https://contoso.sharepoint.com/sites/MyTeamSite', you will have a Client ID/Secret for __MyTeamSite__. If you also require access to files at 'https://contoso.sharepoint.com/sites/MyOtherTeamSite' you will also need to set up a separate Client ID/Secret for __MyOtherTeam__ site.
<br>