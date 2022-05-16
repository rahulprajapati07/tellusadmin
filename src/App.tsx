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
//import UnAuthorizeduser from "../src/component/UnAuthorizedUser"

//import { ProfileContentBackendService } from './component/BackendService';
// import {canUserRestoreTeams} from './component/graph';


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

  readonly authTokenRequest: microsoftTeams.authentication.AuthTokenRequest = {
    successCallback: function (token: string) {
      //const decoded: { [key: string]: any; } = jwt.decode(token);
      //localStorage.setItem("name", decoded.name);
      //localStorage.setItem("token", token);
      console.log("Teams Token :- ", token);
    },
    failureCallback: function (error: any) {
      console.log("Failure on getAuthToken: " + error);
    }
  };

  constructor(props:any){
    super(props);
    microsoftTeams.initialize(() => {
      microsoftTeams.getContext((r) => {
        microsoftTeams.authentication.getAuthToken(this.authTokenRequest);
      });
    });
    this.state = {
      instance : undefined,
      accounts : undefined,
      userIsAdmin : false,
      name : '', 
    }
  }

  componentDidMount(){
    
  }

  render(){

    let accessToken : string = '' ; 

    return (
      <div className="App">
       Token :-  { accessToken }
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