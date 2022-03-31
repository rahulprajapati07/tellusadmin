import React , {Component } from 'react'; //, {useState}
import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { loginRequest } from "./component/authConfig";
import  Button from "react-bootstrap/Button";
//import  DetailsListDemo  from './DataListDemo';
import WorkspaceDetails from './component/Workspace';

//import { ProfileContentBackendService } from './component/BackendService';
// import {canUserRestoreTeams} from './component/graph';


// let userIsAdmin = false;

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
        <WorkspaceDetails instance = {instance} accounts = {accounts}  />
        {/* {
          userIsAdmin == true ?  : <Button variant="secondary" onClick={GetUser}>Get USer </Button> 
        } */}
      </>
    );
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

class App extends Component {
  
  
  // public async handleLogin(instance :any) : Promise<any> {
  //   instance.loginPopup(loginRequest).then((response : any) => {
  //     console.log("Login response",response);
  //   }).catch((e :any)  => {
  //       console.error(e);
  //   });
  //   return new Promise<any>(() => {})
  // }
  // public handleLogin(instance: any) {
  //   instance.loginPopup(loginRequest).catch((e :any) => {
  //       console.error(e);
  //   });
  // }


  // public ProfileContent = () => {
  
  //   const { instance, accounts } = useMsal();
  //     return(
  //       <>
  //         {/* <Button variant="secondary" onClick={RequestAllTeams}>Get All Teams</Button> */}
  //         {/* <Button variant="secondary" onClick={GetPublicTeams}>Get Public Teams</Button> */}
  //         <div>
  //             <DetailsListDemo instance = {instance} accounts = {accounts}  />
              
  //         </div>
  //       </>
  //     );
  // }

  // public MainContent = ()  => {    
  //   const { instance } = useMsal();
  //   // if(instance.getAllAccounts()[0] === undefined)
  //   // {
  //   //    this.handleLogin(instance);
  //   // }
  //   return (
  //       <div className="App">
  //         <div>
  //           Template
  //         </div>
  //           <AuthenticatedTemplate>
  //             <div>
  //               AuthenticatedTemplate 
  //             </div>
  //                <this.ProfileContent />
  //           </AuthenticatedTemplate>
  
  //           <UnauthenticatedTemplate>
  //           <Button variant="secondary" className="ml-auto" onClick={() => this.handleLogin(instance)}>Sign in using Popup</Button>
  //           </UnauthenticatedTemplate>
  //       </div>
  //   );
  // };
  componentDidMount(){
  }

  render(){
    return (
      <div className="App">
        <MainContent />
      </div>
    )
  }
}

export default App;