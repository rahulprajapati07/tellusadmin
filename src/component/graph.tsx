import { graphConfig } from "./authConfig";
import { getAllGroups, getTeams,getPublicTeams } from "./authConfig";

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


