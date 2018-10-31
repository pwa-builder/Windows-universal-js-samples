/**
 * @file
 * MSFT Graph User Authentication
 */

/**
 * This allows you to authenticate your users with a microsoft account or Acitve Directory account to get access to the graph
 *
//  * @alias Create MSFT Graph Authentication 
//  * @method createActivity 
 * @param {object} userAgentApplication The userAgentApplication object will help you doing the authentication job And get the token to do Graph API calls
 * @param {object} user User object return by MSAL lib
 * @param {object} [msalconfig = {
            clientID: "54a0ef8a-3934-49f4-85c8-f8e32511fd93",
            redirectUri: location.origin
        };] Register your app there: https://apps.dev.microsoft.com/portal/register-app & add a web platform to get a Client ID If you already did, retrieve the Client ID from: https://apps.dev.microsoft.com/#/appList
 * @param {object} [graphAPIScopes = ["https://graph.microsoft.com/contacts.read", "https://graph.microsoft.com/user.read", "https://graph.microsoft.com/sites.readwrite.all"]] Permissions you're requesting to do your future Graph API calls
 * @see
 */

        // The userAgentApplication object will help you doing the authentication job
        // And get the token to do Graph API calls
        userAgentApplication = new Msal.UserAgentApplication(msalconfig.clientID, null, loginCallback, {
            redirectUri: msalconfig.redirectUri
        });

        //Previous version of msal uses redirect url via a property
        if (userAgentApplication.redirectUri) {
            userAgentApplication.redirectUri = msalconfig.redirectUri;
        }

        // If page is refreshed, continue to display user info
        if (!userAgentApplication.isCallback(window.location.hash) && window.parent === window && !window.opener) {
            user = userAgentApplication.getUser();
            if (user) {
                console.log("user: " + user.name);
            }
        }

        function showError(endpoint, error, errorDesc) {
            var formattedError = JSON.stringify(error, null, 4);
            if (formattedError.length < 3) {
                formattedError = error;
            }
            console.error(error);
        }

        function loginCallback(errorDesc, token, error, tokenType) {
            if (errorDesc) {
                showError(msal.authority, error, errorDesc);
            } else {
                console.log("You can now do calls to Graph API starting from here.");
            }
        }
        //example element to attache a login button to
        document.getElementById("Login").addEventListener("click", () => {
            // Call this code on the click event of your login button
            userAgentApplication.loginRedirect(graphAPIScopes);   
        });

