import { getPublicTeams, canUserRestoreTeam , archiveTeam } from "./authConfig"; //getAllGroups, getTeams,
import { deleteWorkspace as deleteWorkspaceAPI } from "./authConfig"
import axios from "axios";
/**
 * Attaches a given access token to a MS Graph API call. Returns information about the user
 * @param accessToken 
 */

export async function callGetPublicTeams(accessToken : string) : Promise<any> {
    
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;
    
    headers.append("Authorization", bearer);
    // const config = {
    //     headers: {'content-type': 'application/x-www-form-urlencoded'}
    // };
    const options = {
        method: "POST",
        headers: headers
    };

    return new Promise<any>((resolve, reject) => {  fetch(getPublicTeams.getPublicTeams, options)
        .then(response => response.json())
        .then((data) => {
            console.log("Graph Call : Response For API ");
            resolve(data);
          })
        .catch(error => console.log(error));
    })
}

export async function canUserRestoreTeams(accessToken : string, userMail:string) : Promise<boolean> {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);
    headers.append("Content-Type","application/json");
    headers.append("API-Key","");


    // const config = {
    //     headers: {'content-type': 'application/x-www-form-urlencoded'}
    // };
    const options = {
        method: "POST",
        headers: headers,
        body : JSON.stringify({
            username: userMail,
            adminDirectoryRoleNames: ["Teams Administrator"]    
        })
    };
    return new Promise<boolean>((resolve, reject) => {
        fetch(canUserRestoreTeam.canUserRestoreTeam, options)
        .then(response => response.json())
        .then((data:boolean) => {
            resolve(data);
            console.log("Teams USer" + data);
        })
        .catch(error => console.log(error));
    })
}

export async function deleteWorkspace(accessToken : string, item:any) : Promise<any> {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);
    headers.append("Content-Type","application/json");
    headers.append("API-Key","");
    const options = {
        method: "DELETE",
        headers: headers,
        body : JSON.stringify({
            id: item.teamsGroupId,
        })
    };
    return new Promise<any>((resolve, reject) => {
        fetch(deleteWorkspaceAPI.deleteWorkspace, options)
        .then(response => resolve(response))
        .catch(error => console.log(error));
    })
}

export async function archiveWorkspace(accessToken : string, item:any) : Promise<any> {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);
    headers.append("Content-Type","application/json");
    headers.append("API-Key","");
    const options = {
        method: "POST",
        headers: headers,
        body : JSON.stringify({
            id: item.teamsGroupId,
            option: item.status === "Archived" ? "unarchive" : "archive"
        })
    };
    return new Promise<any>((resolve, reject) => {
        fetch(archiveTeam.archiveTeam, options)
        .then(response => {
            console.log("Archived API Response");
            console.log(response);
            resolve(response);
    })
        .catch(error => console.log(error));
    })
}

export async function getClientDetails(token: any, currentUserEmail: string,  tenantId: any) {
    axios.defaults.headers.post['Content-Type'] = 'application/json';
    let model = {
        TeamsAuthToken: token,
        TenantId: tenantId,
        CurrentUser:currentUserEmail
    };
    return new Promise((resolve, reject) => {
        axios({
            method: 'post',
            url: `https://tellusaccesstoken.azurewebsites.net/api/Token/GetO365Token/`,
            data: model
        }).then(function (response) {
            if (response && response['status'] === 200) {
                let token = response['data']['access_token'];
                resolve(token);
            }
            else {
                reject(null);
            }
        }).catch(function (ex) {
            console.error(ex);
            reject(ex);
        });
    })
}
