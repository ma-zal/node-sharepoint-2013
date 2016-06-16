/*global module*/
"use strict";


/**
 * @author Martin Zaloudek
 * @version 16.06.16
 * @module services/sharepoint-list
 */


/**
 * Module for REST API communication of Shrarepoint 2013 in cloud.
 * Can query and get Sharepoint lists.
 */
module.exports = {
    findAllLists: findAllLists,
    getList: getList,
    getListWithAttachmentFiles : getListWithAttachmentFiles,
    getListItemAttachmentFileUrl: getListItemAttachmentFileUrl
};


// Includes
var Promise = require('bluebird');
var rp = require('request-promise');


/**
 * Get array of sharepoint lists.
 * Where are lists?: Sharepoint site -> lists
 * 
 * @public
 * @param {string} sharepointResourceUrl - Example: "https://sharepoint.mycompany.com".
 * @param {string} siteName - Name of Sharepoint site.
 * @param {string} accessToken - Access token used as Bearer authorization in Shrepoint REST API requests
 * @returns {Promise.<Array.<SharepointList>>}
 */
function findAllLists(sharepointResourceUrl, siteName, accessToken) {
    var apiUrl = `${sharepointResourceUrl}/sites/${siteName}/_api/Web/Lists`;
    return getArrayRecursively(apiUrl,accessToken);
}


/**
 * Get items of specified list.
 * Where are items?: Sharepoint site -> list -> items
 * 
 * @public
 * @param {string} sharepointResourceUrl - Example: "https://sharepoint.mycompany.com".
 * @param {string} siteName - Name of Sharepoint site.
 * @param {string} listGuid - GUID of required list items.
 * @param {string} accessToken - Access token used as Bearer authorization in Shrepoint REST API requests
 * @returns {Promise.<Array.<SharepointListItem>>} - Rows of list
 */
function getList(sharepointResourceUrl, siteName, listGuid, accessToken) {
    var apiUrl = `${sharepointResourceUrl}/sites/${siteName}/_api/Web/Lists(guid'${listGuid}')/Items`;
    return getArrayRecursively(apiUrl, accessToken);
}


/**
 * Get attachment files.
 * Where are files?: Sharepoint site -> list -> item -> attachment files
 * 
 * @public
 * @param {string} sharepointResourceUrl - Example: "https://sharepoint.mycompany.com".
 * @param {string} siteName - Name of Sharepoint site.
 * @param {string} listGuid - GUID of required list items.
 * @param {number} listItemId - Index of item in list.
 * @param {string} accessToken - Access token used as Bearer authorization in Shrepoint REST API requests
 * @returns {Promise.<Array>} - Rows of list
 */
function getListAttachmentFiles(sharepointResourceUrl, siteName, listGuid, listItemId, accessToken) {
    var apiUrl = `${sharepointResourceUrl}/sites/${siteName}/_api/Web/Lists(guid'${listGuid}')/Items(${listItemId})/AttachmentFiles`; 
    return getArrayRecursively(apiUrl, accessToken);
}


/**
 * Get items of specified list. Loads attached files in list items.
 * Structure of items: Sharepoint site -> lists -> items -> attached files.
 * 
 * @public
 * @param {string} sharepointResourceUrl - Example: "https://sharepoint.mycompany.com".
 * @param {string} siteName - Name of Sharepoint site.
 * @param {string} listGuid - GUID of required list items.
 * @param {string} accessToken - Access token used as Bearer authorization in Shrepoint REST API requests
 * @returns {Promise.<Array.<SharepointListItem>>} - Rows of list
 */
function getListWithAttachmentFiles(sharepointResourceUrl, siteName, listGuid, accessToken) {
    return getList(sharepointResourceUrl, siteName, listGuid, accessToken).then(function(/*Array.<SharepointListItem>*/ sharepointList) {
        var subPromises = sharepointList.map(function(/*SharepointListItem*/ sharepointListItem) {

            // Load attachment files
            return getListAttachmentFiles(sharepointResourceUrl, siteName, listGuid, sharepointListItem.Id, accessToken).then(function(attachmentFiles) {
                sharepointListItem.AttachmentFiles = attachmentFiles;
            });

        });
        
        // After all attachments load, return whole structure.
        return Promise.all(subPromises).then(function() {
            return sharepointList;
        });
    })
}


/**
 * Read list of object from REST API.
 * Result can contains `__next` property, that contains URL destination of next values (pagination).
 * If `__next` is present, this metods recursively loads whole pages. 
 * 
 * @private
 * @param {string} apiUrl - Absolute URL of REST API
 * @param {string} accessToken - Access token used as Bearer authorization in Shrepoint REST API requests
 * @returns {Promise.<Array>}
 */
function getArrayRecursively(apiUrl, accessToken) {
    return rp({
        uri: apiUrl,
        headers: {
            Accept: 'application/json;odata=verbose',
            Authorization: 'Bearer ' + accessToken
        },
        strictSSL: false, // Without it sometimes fails with: "write EPROTO 101057795:error:1408F10B:SSL routines:SSL3_GET_RECORD:wrong version number:openssl\ssl\s3_pkt.c:362"
        json: true // Automatically parses the JSON string in the response 

    }).then(function(/*Object*/ apiResponse) {
        // Save returned part of rows
        var promises = [Promise.resolve(apiResponse.d.results)];
        
        if (apiResponse.d.__next) {
            // Recursively load next part of rows
            promises.push(getArrayRecursively(apiResponse.d.__next, accessToken));
        }
        return Promise.all(promises);

    }).then(function(/*Array.<Array>*/ responses) {

        // Merge all parts of rows together
        return responses.reduce(function(/*Array*/ finalArray, /*Array*/ part) {
            return finalArray.concat(part);
        }, []);
    });
    
}

/**
 * 
 * @param sharepointResourceUrl
 * @param siteName
 * @param listGuid
 * @param {number} listItemId - First index is 1 (not 0).
 * @param attachmentFilename
 * @returns {string}
 */
function getListItemAttachmentFileUrl(sharepointResourceUrl, siteName, listGuid, listItemId, attachmentFilename) {
    return `${sharepointResourceUrl}/sites/${siteName}/_api/Web/Lists(guid'${listGuid}')/Items(${listItemId})/AttachmentFiles('${attachmentFilename}')/$value`; 
}


/**
 * @typedef {Object} SharepointListItem
 * @property {string} [__metadata.id] - Some GUID (another value then GUID property)
 * @property {string} [__metadata.uri] - API for get details of this Item.
 * @property {number} Id
 * @property {string} Title
 * @property {string} Description
 * @property {string} [Modified] - Example: '2016-05-24T11:49:03Z'
 * @property {string} [Created] - Example: '2016-05-24T11:49:03Z'
 * @property {Date} ModifiedDate - In filtered instance it is converted value from `Modified` string property.
 * @property {Date} CreatedDate - In filtered instance it is converted value from `Created` string property.
 * @property {string} GUID - Another GUID.
 * @property {Array.<SharepointListItemAttachmentFile>} [AttachmentFiles]
 */

/**
 * @typedef {Object} SharepointListItemAttachmentFile
 * @property {string} FileName - Example: 'week19.pdf'
 * @property {string} ServerRelativeUrl - Example: '/sites/iPadDep/Lists/test for peter/Attachments/3/week19.pdf'
 */