/*global module*/
"use strict";

/**
 * @author Martin Zaloudek
 * @version 16.06.16
 * @module services/sharepoint-files
 */


/**
 * Module for REST API communication of Shrarepoint 2013 in cloud.
 * Can query and get folder and files.
 * 
 * @see {@link https://msdn.microsoft.com/en-us/library/office/dn292553.aspx}
 */
module.exports = {
    getSubfolders: getSubfolders,
    getFiles: getFiles
};


// Includes
var rp = require('request-promise');


/**
 * Get subfolders of specified directory.
 * 
 * @public
 * @param {string} sharepointResourceUrl - Example: "https://sharepoint.mycompany.com".
 * @param {string} siteName - Name of Sharepoint site.
 * @param {string} folderPath - Path of directory, that you want to query.
 * @param {string} accessToken - Access token used as Bearer authorization in Shrepoint REST API requests
 * @returns {Promise.<Array.<FolderMedatada>>}
 */
function getSubfolders(sharepointResourceUrl, siteName, folderPath, accessToken) {
    var _folderPath = (folderPath.indexOf("/sites/" + siteName) === 0 ? "" : "/sites/" + siteName + "/") + folderPath;
    var apiUrl = sharepointResourceUrl + "/sites/" + siteName + "/_api/Web/GetFolderByServerRelativeUrl('" + _folderPath + "')/Folders";
    return callApi(apiUrl, accessToken);
}


/**
 * Get files of specified directory.
 * 
 * @public
 * @param {string} sharepointResourceUrl - Example: "https://sharepoint.mycompany.com".
 * @param {string} siteName - Name of Sharepoint site.
 * @param {string} folderPath - Path of directory, that you want to query.
 * @param {string} accessToken - Access token used as Bearer authorization in Shrepoint REST API requests
 * @returns {Promise.<Array.<FileMedatada>>}
 */
function getFiles(sharepointResourceUrl, siteName, folderPath, accessToken) {
    var _folderPath = (folderPath.indexOf("/sites/" + siteName) === 0 ? "" : "/sites/" + siteName + "/") + folderPath;
    var apiUrl = sharepointResourceUrl + "/sites/" + siteName + "/_api/Web/GetFolderByServerRelativeUrl('" + _folderPath + "')/Files";
    return callApi(apiUrl, accessToken);
}


/**
 * Call Sharepoint REST API.
 * 
 * @private
 * @param {string} apiUrl - Absolute path of API call
 * @param {string} accessToken - Access token used as Bearer authorization in Shrepoint REST API requests
 * @returns {Promise.<Array.<Object>>}
 */
function callApi(apiUrl, accessToken) {
    return rp({
        uri: apiUrl,
        headers: {
            Accept: 'application/json;odata=verbose',
            Authorization: 'Bearer ' + accessToken
        },
        strictSSL: false, // Without it sometimes fails with: "write EPROTO 101057795:error:1408F10B:SSL routines:SSL3_GET_RECORD:wrong version number:openssl\ssl\s3_pkt.c:362"
        json: true // Automatically parses the JSON string in the response 

    }).then(function(/*Object*/ apiResponse) {
        return apiResponse.d.results;
    });
}


/**
 * @typedef {Object} FolderMedatada
 * @property {string} UniqueId          - Example: '892a8ca7-7f5e-400b-b426-78c7384ea5bd'
 * @property {string} ServerRelativeUrl - Example: '/sites/iPadDep/Style Library/Media Player'
 * @property {string} Name              - Firectory name. Example: 'Media Player'
 * @property {number} ItemCount
 * @property {string} TimeCreated       - Example: '2015-08-20T18:22:50Z'
 * @property {string} TimeLastModified  - Example: '2015-08-20T18:22:50Z',
 */

/**
 * @typedef {Object} FileMedatada
 * @property {string} Name              - File name. Example: 'AlternateMediaPlayer.xaml'
 * @property {string} ServerRelativeUrl - Example: '/sites/iPadDep/Style Library/Media Player/AlternateMediaPlayer.xaml'
 * @property {number} MajorVersion
 * @property {number} MinorVersion
 * @property {string} Length            - It is number in bytes as string. Example: '7110'.
 * @property {string|null} Title
 * @property {string} TimeCreated       - Example: '2015-08-20T18:22:50Z'
 * @property {string} TimeLastModified  - Example: '2015-08-20T18:22:50Z',
 * @property {string} UniqueId          - Example: '8e71fe00-f9ee-4493-b9dc-e95df1802b43'
 */
