// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

const getAccessToken = async function () {
    console.log("getAccessToken called with arguments:", arguments);
    
    // Create a config variable that store credentials from config.json
    const config = require(__dirname + "/../config/config.local.json");

    // Use MSAL.js for authentication
    const msal = require("@azure/msal-node");

    // Print out the configuration being used
    console.log("Authentication configuration:");
    console.log("- authenticationMode:", config.authenticationMode);
    console.log("- clientId:", config.clientId);
    console.log("- tenantId:", config.tenantId);
    
    const msalConfig = {
        auth: {
            clientId: config.clientId,
            authority: `${config.authorityUrl}${config.tenantId}`,
        }
    };

    // Check for the MasterUser Authentication
    if (config.authenticationMode.toLowerCase() === "masteruser") {
        const clientApplication = new msal.PublicClientApplication(msalConfig);

        const usernamePasswordRequest = {
            scopes: [config.scopeBase],
            username: config.pbiUsername,
            password: config.pbiPassword
        };
        
        console.log("Using MasterUser authentication with:", { 
            scopes: [config.scopeBase],
            username: config.pbiUsername,
            // Not logging password for security
        });

        return clientApplication.acquireTokenByUsernamePassword(usernamePasswordRequest);

    };

    // Service Principal auth is the recommended by Microsoft to achieve App Owns Data Power BI embedding
    if (config.authenticationMode.toLowerCase() === "serviceprincipal") {
        msalConfig.auth.clientSecret =  config.clientSecret
        const clientApplication = new msal.ConfidentialClientApplication(msalConfig);

        const clientCredentialRequest = {
            scopes: [
                config.scopeBase,
                // 'https://analysis.windows.net/powerbi/api/Report.ReadWrite.All',
                // 'https://analysis.windows.net/powerbi/api/Dataset.ReadWrite.All'
            ],
        };
        
        console.log("Using ServicePrincipal authentication with scopes:", [clientCredentialRequest.scopes]);

        const result = await clientApplication.acquireTokenByClientCredential(clientCredentialRequest);
        console.log("acquireTokenByClientCredential_result: ", result);
        return result;
    }
}

module.exports.getAccessToken = getAccessToken;