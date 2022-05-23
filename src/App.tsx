import React  from 'react'; //, {useState}
import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { loginRequest } from "./component/authConfig";
//import { useJwt } from "react-jwt";
//import  Button from "react-bootstrap/Button";
//import  DetailsListDemo  from './DataListDemo';
import WorkspaceDetails from './component/Workspace';
//import { promises } from 'fs';
import {canUserRestoreTeams}  from "../src/component/graph";
import * as microsoftTeams from "@microsoft/teams-js";
//import jwtDecode from "jwt-decode";
//import UnAuthorizeduser from "../src/component/UnAuthorizedUser"

//import { ProfileContentBackendService } from './component/BackendService';
import {getClientDetails} from './component/graph';


// let userIsAdmin = false;

let checkuserIsAdmin : any;

function handleLogin(instance :any,accounts:any) {
    instance.loginPopup(loginRequest).catch((e :any)  => {
        console.error(e);
    });
}
const ProfileContent = () => {
  const { instance, accounts } = useMsal();
      return (
        <> 
            <WorkspaceDetails instance = {instance} accounts = {accounts} userIsAdmin = {checkuserIsAdmin}  />
        </>
        )
}
/**
 * If a user is authenticated the ProfileContent component above is rendered. Otherwise a message indicating a user is not authenticated is rendered.
 */
const MainContent = () => {
  const { instance , accounts } = useMsal();
  
  var loginSuccess = 1;
  for (let index = 0; index <= loginSuccess; index++) {
    if(instance.getAllAccounts()[0] === undefined)
        {
          handleLogin(instance,accounts);
        }
  }

  checkUserRole();

  function checkUserRole() {
    instance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0]
    }).then((response : any) => 
    {
      canUserRestoreTeams(response.accessToken, accounts[0].username).then(response => response ).then( (data:any) =>
      {
        checkuserIsAdmin = data;
      })
    })
  }

  return (
      <div className="App">

          <AuthenticatedTemplate>
               <ProfileContent />
          </AuthenticatedTemplate>

          <UnauthenticatedTemplate>
          {/* <Button variant="secondary" className="ml-auto" onClick={() => handleLogin(instance,accounts)}>Sign in using Popup</Button> */}
          </UnauthenticatedTemplate>
      </div>
  );
};

// function getAccount () {

// }
// const { instance , accounts } = useMsal();
class App extends React.Component {

  constructor(props:any){
    super(props);
    microsoftTeams.initialize();
    this.state = {
      instance : undefined,
      accounts : undefined,
      userIsAdmin : false,
      name : '', 
    }
  }

  componentDidMount(){

    microsoftTeams.authentication.getAuthToken({
      successCallback: (token: string) => {
          //const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
          //setName(decoded!.name);
          microsoftTeams.appInitialization.notifySuccess();

          getClientDetails(token + "", "belinda@iiab.onmicrosoft.com", "082a7423-5b17-4f5e-a4dc-6d2396d7edfa").then((graphToken) => {
              console.log(graphToken);
          }).catch((err) => {
              console.log(err);
          })

      },
      failureCallback: (message: string) => {
          //setError(message);
          microsoftTeams.appInitialization.notifyFailure({
              reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
              message
          });
      },
      resources:["https://antaresbots.onmicrosoft.com/tellus-dev-api"]
  });



    // microsoftTeams.authentication.getAuthToken({
    //   successCallback: (token: string) => {


    //     console.log("Access Token :- ", token);
    //   },

    //   failureCallback: (message: string) => {
    //       console.log("Teams Error FailureCallback :- " ,message);
    //   },
    //   resources:["api://ambitious-pebble-0b2637f10.1.azurestaticapps.net/b0785c01-bd69-4a12-bfe1-e558e7a4b7d1"]
    // });
  }

  render(){

 //   let accessToken : string = '' ; 

    return (
      <div className="App">
        <MainContent />
        {/* <AuthenticatedTemplate>
               <ProfileContent />
          </AuthenticatedTemplate> */}
      </div>
    )
  }
}

// interface ILoginConfig {
//   instance : any,
//   accounts : any,
//   userIsAdmin : boolean,
//   name : string
// }

export default App;