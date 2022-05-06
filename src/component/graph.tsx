import { graphConfig } from "./authConfig";
import { getAllGroups, getTeams,getPublicTeams, canUserRestoreTeam , archiveTeam } from "./authConfig";
import { deleteWorkspace as deleteWorkspaceAPI } from "../component/authConfig"
/**
 * Attaches a given access token to a MS Graph API call. Returns information about the user
 * @param accessToken 
 */
export async function callMsGraph(accessToken : string) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    return fetch(graphConfig.graphMeEndpoint, options)
        .then(response => response.json())
        .catch(error => console.log(error));
}

export async function callMsGraphGroup(accessToken : string) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    return fetch(getAllGroups.getGroups, options)
        .then(response => response.json())
        .catch(error => console.log(error));
}

export async function callAllTeamsRequest(accessToken : string) : Promise<any> {
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

    return new Promise<any>((resolve, reject) => {
        fetch(getTeams.getAllTeams, options)
        .then(response => response.json())
        .then((data) => {
            resolve(data);
        })
        .catch(error => console.log(error));
    }) 
}

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
