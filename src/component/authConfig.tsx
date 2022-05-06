export const msalConfig = {
    auth: {
      clientId: "b0785c01-bd69-4a12-bfe1-e558e7a4b7d1",
      authority: "https://login.microsoftonline.com/common", // This is a URL (e.g. https://login.microsoftonline.com/{your tenant ID})
      //redirectUri: "http://localhost:3000/",
      //redirectUri: window.location.origin,
      redirectUri: "https://ambitious-pebble-0b2637f10.1.azurestaticapps.net/"
    },
    cache: {
      cacheLocation: "sessionStorage", // This configures where your cache will be stored
      storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    }
  };
  
  // Add scopes here for ID token to be used at Microsoft identity platform endpoints.
  export const loginRequest = {
   scopes: ["https://antaresbots.onmicrosoft.com/tellus-dev-api/user_impersonation"]
  };
  
/**
 * Add here the scopes to request when obtaining an access token for MS Graph API. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */

export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com"
};

export const getAllGroups = {
  getGroups: "https://graph.microsoft.com/v1.0/groups/"
};

export const getTeams = {
  getAllTeams : "https://tellus-dev-api.azurewebsites.net/api/GetMyTeams?"
};

// export const getPublicTeams = {
//   getPublicTeams : "https://ffde-40-88-125-34.ngrok.io/api/GetPublicTeams?"
// }
export const getPublicTeams = {
  getPublicTeams : "https://tellus-dev-api.azurewebsites.net/api/GetMyTeams?"
};

export const deleteWorkspace = {
  deleteWorkspace : "https://tellus-dev-api.azurewebsites.net/api/DeleteGroup?"
};

export const canUserRestoreTeam = {
  canUserRestoreTeam : "https://tellus-dev-api.azurewebsites.net/api/CanUserRestoreTeams?"
};

export const archiveTeam = {
  archiveTeam : "https://tellus-dev-api.azurewebsites.net/api/ArchiveTeam?"
}