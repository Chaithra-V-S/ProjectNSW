import * as React from 'react';
import styles from './NswNewEpisodeOfCare.module.scss';
import { INswNewEpisodeOfCareProps } from './INswNewEpisodeOfCareProps';
import { INswNewEpisodeOfCareState } from './INswNewEpisodeOfCareState';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { Text } from 'office-ui-fabric-react/lib/Text';
import {Button, ButtonType} from 'office-ui-fabric-react/lib/Button';
import { MSGraphClient } from "@microsoft/sp-http";

export default class NswNewEpisodeOfCare extends React.Component<INswNewEpisodeOfCareProps, INswNewEpisodeOfCareState> {
  
  constructor(props: INswNewEpisodeOfCareProps) {
    super(props);
    this.state={
      siteName:"",
      siteDescription:""
    }
    this.createteam = this.createteam.bind(this);
  }
  public createteam()
  {
    let createTeamSite:any={
    
      "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      "displayName": this.state.siteName,
      "description": this.state.siteDescription   
  };
    this.props.context.msGraphClientFactory
        .getClient() // Init Microsoft Graph Client
        .then((client: MSGraphClient): void => {
            client
                .api(`/teams`) //Get teams method
                .version("v1.0") 
                .post(createTeamSite)
                .then((response)=>{
                  alert("teams created")  ;
                  this.setState({siteName:"",siteDescription:""});
                });
        });
  }
  public render(): React.ReactElement<INswNewEpisodeOfCareProps> {
    return (
      <div className={ styles.nswNewEpisodeOfCare }>
        <TextField
          label="Enter Teams Name"
          value={this.state.siteName} 
          onChanged={e => {this.setState({siteName:e})} }
        />
         <TextField
        label="Enter teams description"
        value={this.state.siteDescription} 
        onChanged={e => {this.setState({siteDescription:e})} }
      />
      <br/>
         <Button buttonType={ButtonType.primary} onClick={this.createteam}>create Episode Of Care</Button>
     
      </div>
    );
  }
}
