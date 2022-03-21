// import React, {useState} from 'react';
// import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import {callAllTeamsRequest} from "./graph"; //{ callMsGraph,callMsGraphGroup,callAllTeamsRequest,callGetPublicTeams } 


export function ProfileContentBackendService  (instance : any,accounts : any)  {
        
        instance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        }).then((response : any) => {
            callAllTeamsRequest(response.accessToken).then(response => response).then((data:any[]) =>
            {
                console.log("Request all Teams data",data);
            });
        });
        // function getAllPublicTeams() {
        //     instance.acquireTokenSilent({
        //         ...loginRequest,
        //         account: accounts[0]
        //     }).then((response) => {
        //         callAllTeamsRequest(response.accessToken).then(response => response).then((data:any[]) =>
        //         {
        //             console.log("Request all Teams data",data);
        //         });;
        //     });
        // }
}



// export async function GetAllPublicTeams() : Promise<any> {
//     getAllPublicTeams();
// }