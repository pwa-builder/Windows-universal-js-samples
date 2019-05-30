/**
 * @file
 * MSFT Graph User Authentication
 */

/**
 * This allows you to authenticate your users with a Microsoft account or Active Directory account to get access to the graph
 *
 * @alias Microsoft Graph Authentication 
 * @method authWithGraph
 * @param {object} [scopes = ""] Array of API URLs you are requesting permissions for future Graph API calls
 * @param {object} [clientID = ""] Follow these docs to register your app and receive a clientID https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-core/README.md#prerequisite
 * @see https://raw.githubusercontent.com/pwa-builder/Windows-universal-js-samples/dev/win10/images/graphAuth.png
 */


const scopes = ["https://graph.microsoft.com/contacts.read", "https://graph.microsoft.com/user.read", "https://graph.microsoft.com/sites.readwrite.all"];

async function authWithGraph(scopes, clientID) {
  if (clientID && scopes) {
    const userAgentApplication = new Msal.UserAgentApplication(config.clientID, null, authRedirectCallback);
    try {
      await userAgentApplication.loginPopup(scopes);
    }
    catch (error) {
      console.error('Error during login', error);
    }

    try {
      // Login success
      const accessToken = await userAgentApplication.acquireTokenSilent(scopes);
      return accessToken;
    }
    catch (error) {
      // AcquireTokenSilent Failure, send an interactive request.
      // This will show the Microsoft Account login UI again
      const accessToken = await userAgentApplication.acquireTokenPopup(scopes)
      return accessToken;
    }
  } else {
    console.log("You must supply a client id and authentication scopes for your app");
  }
}


function authRedirectCallback(errorDesc, token, error, tokenType) {
  if (error) {
    console.error(errorDesc, error);
  } else {
    return token;
  }
}
 
