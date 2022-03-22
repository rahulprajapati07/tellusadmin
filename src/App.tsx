import React,{ Component } from 'react'; //, {useState}
import './App.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { loginRequest } from "./component/authConfig";
import Button from "react-bootstrap/Button";
import  { DetailsListDemo } from './DataListDemo';
//import { ProfileContentBackendService } from './component/BackendService';



  // function handleLogin(instance :any) {
  //   instance.loginRedirect(loginRequest).catch((e :any)  => {
  //       console.error(e);
  //   });
  // }

// const ProfileContent = () => {
  
//   const { instance, accounts } = useMsal();
//   //const [graphData, setGraphData] = useState(null);

// //   function RequestAllTeams() {
// //     // Silently acquires an access token which is then attached to a request for MS Graph data
// //     ProfileContentBackendService(instance,accounts);
// //     // GetAllPublicTeams().then((data:any[]) =>
// //     // {
// //     //     console.log("Request all Teams data 1 :",data);
// //     // });
// //     // instance.acquireTokenSilent({
// //     //     ...loginRequest,
// //     //     account: accounts[0]
// //     // }).then((response) => {
// //     //     callAllTeamsRequest(response.accessToken).then(response => response).then((data:any[]) =>
// //     //     {
// //     //         console.log("Request all Teams data",data);
// //     //     });;
// //     // });
// // }

// // function GetPublicTeams() {
// //     // Silently acquires an access token which is then attached to a request for MS Graph data
// //     instance.acquireTokenSilent({
// //         ...loginRequest,
// //         account: accounts[0]
// //     }).then((response) => {
// //         callGetPublicTeams(response.accessToken).then(response => response).then((data:any[]) =>
// //         {
// //           console.log("Get public Teams data",data);
// //         });
// //     });
// // }

//     return(
//       <>
//         {/* <Button variant="secondary" onClick={RequestAllTeams}>Get All Teams</Button> */}
//         {/* <Button variant="secondary" onClick={GetPublicTeams}>Get Public Teams</Button> */}
//         <div>
//             <DetailsListDemo instance = {instance} accounts = {accounts}  />
//             {/* <div className={ styles.searchBoxContainer}>
//                   <SearchBox 
//                       value="Search box"
//                       className={styles.searchBoxUser}
//                   />
//             </div> */}
//         </div>
//       </>
//     );
// }

/**
 * If a user is authenticated the ProfileContent component above is rendered. Otherwise a message indicating a user is not authenticated is rendered.
 */
// const MainContent = () => {    
//   const { instance } = useMsal();
//   if(instance.getAllAccounts() == null)
//   {
//     handleLogin(instance);
//   }
//   return (
//       <div className="App">
//         <ProfileContent />
//           <AuthenticatedTemplate>
//                <ProfileContent />
//           </AuthenticatedTemplate>

//           {/* <UnauthenticatedTemplate>
//           <Button variant="secondary" className="ml-auto" onClick={() => handleLogin(instance)}>Sign in using Popup</Button>
//           </UnauthenticatedTemplate> */}
//       </div>
//   );
// };

class App extends Component {
  

  public async handleLogin(instance :any) : Promise<any> {
    instance.loginPopup(loginRequest).then((response : any) => {
      console.log("Login response",response);
    }).catch((e :any)  => {
        console.error(e);
    });
    return new Promise<any>(() => {})
  }

  public ProfileContent = () => {
  
    const { instance, accounts } = useMsal();
      return(
        <>
          {/* <Button variant="secondary" onClick={RequestAllTeams}>Get All Teams</Button> */}
          {/* <Button variant="secondary" onClick={GetPublicTeams}>Get Public Teams</Button> */}
          <div>
              <DetailsListDemo instance = {instance} accounts = {accounts}  />
              
          </div>
        </>
      );
  }

  public MainContent = ()  => {    
    const { instance } = useMsal();
    // if(instance.getAllAccounts()[0] === undefined)
    // {
    //    this.handleLogin(instance);
    // }
    return (
        <div className="App">
            <AuthenticatedTemplate>
                 <this.ProfileContent />
            </AuthenticatedTemplate>
  
            <UnauthenticatedTemplate>
            <Button variant="secondary" className="ml-auto" onClick={() => this.handleLogin(instance)}>Sign in using Popup</Button>
            </UnauthenticatedTemplate>
        </div>
    );
  };
  // componentDidMount(){
  //   SignOutButton();
  // }

  render(){
    return (
      <div className="App">
      <this.MainContent />
    </div>
    )
  }
}

export default App;