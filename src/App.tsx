import React  from 'react'; //, {useState}
import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { loginRequest } from "./component/authConfig";
import  Button from "react-bootstrap/Button";
//import  DetailsListDemo  from './DataListDemo';
import WorkspaceDetails from './component/Workspace';
//import { promises } from 'fs';
import {canUserRestoreTeams}  from "../src/component/graph";
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

  // function CheckUserAdmin(instance:any,accounts:any) {
  //         instance.acquireTokenSilent({
  //           ...loginRequest,
  //           account: accounts[0]
  //       }).then((response:any)  => {
  //         canUserRestoreTeams(response.accessToken,accounts[0].username).then((response :boolean )=> response).then((data:boolean) =>
  //           {
  //             console.log("UserAdmin 0 :- " + userIsAdmin);
  //             if(data == true){
  //               userIsAdmin = true;
  //               console.log("UserAdmin :- " + userIsAdmin);
  //             }
  //             console.log("userIsAdmin status :" + userIsAdmin);
  //           });
  //       });
  // }

const ProfileContent = () => {

  const { instance, accounts } = useMsal();
  
  

  //CheckUserAdmin(instance, accounts);
  //const [graphData, setGraphData] = useState(null);

//   function RequestAllTeams() {
//     // Silently acquires an access token which is then attached to a request for MS Graph data
//     ProfileContentBackendService(instance,accounts);
//     // GetAllPublicTeams().then((data:any[]) =>
//     // {
//     //     console.log("Request all Teams data 1 :",data);
//     // });
//     // instance.acquireTokenSilent({
//     //     ...loginRequest,
//     //     account: accounts[0]
//     // }).then((response) => {
//     //     callAllTeamsRequest(response.accessToken).then(response => response).then((data:any[]) =>
//     //     {
//     //         console.log("Request all Teams data",data);
//     //     });;
//     // });
// }

// function GetUser() {
//     // Silently acquires an access token which is then attached to a request for MS Graph data

//     instance.acquireTokenSilent({
//         ...loginRequest,
//         account: accounts[0]
//     }).then((response) => {
//       canUserRestoreTeams(response.accessToken,accounts[0].username).then((response :boolean )=> response).then((data:boolean) =>
//         {
//           console.log("UserAdmin 0 :- " + userIsAdmin);
//           if(data == true){
//             userIsAdmin = true;
//             console.log("UserAdmin :- " + userIsAdmin);
//           }
//           console.log("userIsAdmin status :" + userIsAdmin);
//         });
//     });
// }
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
  //CheckUserAdmin();
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
          <Button variant="secondary" className="ml-auto" onClick={() => handleLogin(instance,accounts)}>Sign in using Popup</Button>
          </UnauthenticatedTemplate>
      </div>
  );
};

// function getAccount () {

// }
// const { instance , accounts } = useMsal();
class App extends React.Component {

  // constructor(props : ILoginConfig){
  //   super(props);
  //   this.checkUserLogin();
  //   this.state = {
  //     instance : undefined,
  //     accounts : undefined,
  //     userIsAdmin : false,
  //   }
  // }

  // public async checkUserLogin(){


  //   for (let index = 0; index <= 1; index++) {
  //     if(instance.getAllAccounts()[0] === undefined)
  //         {
  //           handleLogin(instance,accounts);
  //         }
  //   }
  // }


  componentDidMount(){
  }

  render(){
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
// }

export default App;