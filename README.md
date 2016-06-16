NodeJS Sharepoint 2013 connector
================================

**This library help you to easily connect to your Sharepoint 2013 from NodeJS application and READ lists and files.**

*The most complicated think in Sharepoint REST API communication is understanding of Sharepoint authorization.
And it is solved by official Microsoft [ADAL-NODE](https://github.com/AzureAD/azure-activedirectory-library-for-nodejs) library.
All other is only about "calling right URL" with "Authorization Bearer Token".*

*This library and examples could help you to better understanding, which REST API URL do what, and what values are returned.*


Features
--------

- **Find all Lists at Sharepoint Site**
- **Get rows and values of list**
- **Get files attached into list row**
- **Browse files uploaded to Sharepoint site (out of list)**

Authorization
-------------
For authorization to Sharepoint there is used [offical Microsoft's ADAL NodeJS library](https://github.com/AzureAD/azure-activedirectory-library-for-nodejs) in this library.

For using Sharepoint 2013 REST API, you must:
 - register your NodeJS application in Sharepoint Server (you will obtain Client ID).
 - know the "Tenants". It is GUID of your Sharepoint instance on server .(Because there can exist many instances of Sharepoints at one server). 
 - create/use some user acount and it's password in authorization


Examples
--------

#### Configuration of connection and library loading

```javascript
/** @type {SharepointSourceConfig} */
var sharepointSourceConfig = {	
    /*
     * Authority server for Sharepoint 2013 in cloud
     */
    sharepointAuthorityHostUrl: 'https://login.microsoftonline.com',

    /*
     * Tenant is UUID of your SharePoint instance in Microsoft cloud
     */
    sharepointTenant: '12345678-0000-0000-0000-someyourguid',

    /*
     * URL location of Sharepoint 2013 instance in cloud.
     */
    sharepointResourceUrl: 'https://mysharepoint.mydomain.com',

    /*
     * Client ID identifies your application.
     * Application under this UUID is registered in Sharepoint server.
     * Permissions for Client ID must be:
     *     read user files, profiles, and read items in all site collections
     */
    sharepointClientId: '00000000-0app-guid-0you-registered00',

    /*
     * System credentials used for login into Sharepoint 2013 in cloud
     */
    sharepointUserId: 'SYS_ACCOUNT_SHAREPOINT@mydomain.com',

    /**
     * System credentials used for login into Sharepoint 2013 in cloud
     */
    sharepointUserSecret: 'MySomePassword'
};

// Library loading
var sharepointTokenService = require('node-sharepoint-2013/lib/sharepoint-token')(sharepointSourceConfig);
var sharepointListService = require('node-sharepoint-2013/lib/sharepoint-lists');
```


#### Enable ADAL debugging / logging

Because ADAL communication is quite complicated, it is good idea to enable verbose log to console, when somethink is not working.

```javascript
sharepointTokenService.enableGlobalAdalLogging();
// Note: `sharepointTokenService` variable is from previous example.
```


#### Get all Sharepoint lists

```javascript
// DO NOT FORGOT
// to use config of sharepointSourceConfig and depencencies requires from the previous example!!!


// CONFIGURATION - PLEASE READ AND EDIT

// You will see your site name in sharepoint URL: https://mysharepoint.mydomain.com/sites/mySite/
var mySharepointSite = 'mySite'

// Use `sharepointSourceConfig` variable from previous example	

// END CONFIGURATION


// Get all lists
sharepointTokenService.getToken().then(function(/*TokenResponse*/ tokenResponse) {
   
    return sharepointListService.findAllLists(
        sharepointSourceConfig.sharepointResourceUrl,
        mySharepointSite,
        tokenResponse.accessToken

    ).then(function(lists) {
        console.log("LISTS on site " + mySharepointSite);
        console.log("---------------------------");
        console.log("GUID BaseTemplate Title");
        console.log("---------------------------");
        lists.map(function(list) {
            console.log(list.Id, list.BaseTemplate, list.Title);
        });
    });
});
```


### Get Sharepoint list details

```javascript
// CONFIGURATION - PLEASE READ AND EDIT

// You can find list GUID in console out of previous example
var sharepointListGUID = '00000000-0000-0000-0000-list0guid000';

// Use `sharepointSourceConfig` variable from previous example	
// Use `mySharepointSite` variable from previous example	

// END CONFIGURATION


// Get list details 
sharepointToken.getToken().then(function(/*TokenResponse*/ tokenResponse) {
    return sharepointListService.getListWithAttachmentFiles(
        sharepointSourceConfig.sharepointResourceUrl,
        mySharepointSite,
        sharepointListGUID,
        tokenResponse.accessToken

    ).then(function(list) {
        console.log("Content of list " + mySharepointSite + " " + sharepointListGUID);
        console.log("---------------------------");
        console.log("");
        list.map(function(item) {
            console.log(item);
            console.log("");
            console.log("---------------------------");
            console.log("");
        });
    });
});
```
