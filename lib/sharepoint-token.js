/*global module, process*/
"use strict";

/**
 * @author Martin Zaloudek
 * @version 16.06.16
 * @module services/sharepoint-token
 */

// Global configuration
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";


// Includes
var AuthenticationContext = require('adal-node').AuthenticationContext;
var adal = require('adal-node/lib/adal');
var Promise = require('bluebird');


/**
 * Module for get new Access Token, that will be used in next REST API communication.
 * 
 * This is a factory, that returns object with token management functions for specified sharepoint server. 
 * 
 * @public
 * @name exports
 * @type {function}
 * @param {SharepointSourceConfig} sharepointConfig
 * @returns {{getToken: function, dropToken: function, enableGlobalAdalLogging: function}}
 */
module.exports = function(sharepointConfig) {
    
    var authenticationContext = Promise.promisifyAll(new AuthenticationContext(
        sharepointConfig.sharepointAuthorityHostUrl + '/' + sharepointConfig.sharepointTenant
    ));
    
    /** @type {Promise.<TokenResponse>|null} */
    var tokenRequestPromise = null;
    
    return {
        getToken: getToken,
        dropToken: dropToken,
        enableGlobalAdalLogging: enableGlobalAdalLogging
    };
    
    
    /**
     * Get AccessToken, that will be used in future authorization.
     * AccessToken is cached. If origin AcessToken is valid (by expiration time), it is returned.
     * When it expires, the new one is required.
     * 
     * @private
     * @returns {Promise.<TokenResponse>}
     */
    function getToken() {
        if (tokenRequestPromise && tokenRequestPromise.isFulfilled()) {
            var tokenExpired = (new Date().getTime()) > (Date.parse(tokenRequestPromise.value().expiresOn) + 30000);
            if (tokenExpired) {
                tokenRequestPromise = null;
            }
        }

        if (!tokenRequestPromise || tokenRequestPromise.isCancelled() || tokenRequestPromise.isRejected()) {
            tokenRequestPromise = authenticationContext.acquireTokenWithUsernamePasswordAsync(
                sharepointConfig.sharepointResourceUrl,
                sharepointConfig.sharepointUserId,
                sharepointConfig.sharepointUserSecret,
                sharepointConfig.sharepointClientId
            );
        }
        
        return tokenRequestPromise;
    }

    /**
     * Drop cached accessToken
     */
    function dropToken() {
        tokenRequestPromise = null;
    }
    
    
    /**
     * Enable detailed logging of ADAL library into console.log
     * 
     * @public
     */
    function enableGlobalAdalLogging() {
        var log = adal.Logging;
        log.setLoggingOptions({
            level: log.LOGGING_LEVEL.VERBOSE,
            log: function (level, message, error) {
                console.log(message);
                if (error) {
                    console.log(error);
                }
            }
        });
    }

};


/**
 * @global
 * @typedef {Object} TokenResponse
 * @property {string} tokenType - "bearer" as default
 * @property {number} expiresIn
 * @property {string} expiresOn - Date in string
 * @property {string} resource
 * @property {string} accessToken
 * @property {boolean} isMRRT
 * @property {string} _clientId
 * @property {string} _authority
 */


/**
 * @global
 * @typedef {Object} SharepointSourceConfig
 * @property {string} sharepointAuthorityHostUrl
 * @property {string} sharepointTenant
 * @property {string} sharepointResourceUrl
 * @property {string} sharepointClientId
 * @property {string} sharepointUserId
 * @property {string} sharepointUserSecret
 */