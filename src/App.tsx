import React  from 'react'; //, {useState}
import './App.css';
import WorkspaceDetails from './component/Workspace';
import * as microsoftTeams from "@microsoft/teams-js";
import {getClientDetails} from './component/BackendService';

class App extends React.Component {

  constructor(props:any){
    super(props);
    microsoftTeams.initialize();
  }

  componentDidMount(){

        microsoftTeams.authentication.getAuthToken({
              successCallback: (token: string) => {
                console.log("Access Token For Teams : " + token);
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
                resources:["api://ambitious-pebble-0b2637f10.1.azurestaticapps.net/b0785c01-bd69-4a12-bfe1-e558e7a4b7d1"]
        });
  }

  render(){
    return (
      <div className="App">
        <WorkspaceDetails userIsAdmin = {true} />
      </div>
    )
  }
}

export default App;