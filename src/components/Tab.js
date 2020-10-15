// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";
import { TeamsCloud } from "teamsauth";

/**
 * The 'PersonalTab' component renders the main tab content
 * of your app.
 */
class Tab extends React.Component {
  constructor(props) {
    super(props)
    this.state = {
      context: {},
      userInfo: {},
      profile: {},
      photoObjectURL: "",
      showLoginBtn: false,
      apiError: false,
      result: "",
    }
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  async componentDidMount() {
    // Get the user context from Teams and set it in the state
    microsoftTeams.getContext((context, error) => {
      this.setState({
        context: context
      });
    });
    // Next steps: Error handling using the error object
    await this.initTeamsCloud();
    await this.callGraphSilent();
  }

  async initTeamsCloud() {
    await TeamsCloud.init("https://teamsruntimeconnectorzhijie.azurewebsites.net");
    var userInfo = TeamsCloud.getUserInfo();
    this.setState({
      userInfo: userInfo
    });
  }

  async callGraphSilent() {
    try {
      var graphClient = await TeamsCloud.getMicrosoftGraphClient();
      var profile = await graphClient.api("/me").get();
      var photoBlob = await graphClient.api("/me/photos('120x120')/$value").get();
      this.setState({
        profile: profile,
        photoObjectURL: URL.createObjectURL(photoBlob),
      });
    }
    catch (err) {
      alert("You need to click login button to consent the access: " + err.message);
      this.setState({
        showLoginBtn: true
      });
    }
  }

  async loginBtnClick() {
    try {
      await TeamsCloud.popupLoginPage();
    }
    catch(err) {
      alert("Login failed: " + err);
      return;
    }

    await this.callGraphSilent();
  }

async getFunction(){
    try{
      fetch("https://zhijiefunc1.azurewebsites.net/api/HttpTrigger1", {
        method: 'Get',
      })
        .then(response => {
          if (response.ok) {
            return response.text();
          }
          else{
            this.setState({
              apiError: true,
              result: "Get Azure Function failed with status: " + response.status,
            });
            return;
          }
        })
        .then(data => {
          this.setState({
            apiError: false,
            result: data,
          });
        });
    }
    catch(err) {
      alert("Function trigger failed: " + err);
      return;
    }
  }

  async nameChanged(e) {
    this.setState({ name: e.target.value });
  }

  async postFunction() {
    try{
      fetch("https://zhijiefunc1.azurewebsites.net/api/HttpTrigger1", {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          name: this.state.name,
        }),
      })
        .then(response => {
          if (response.ok) {
            return response.text();
          }
          else{
            alert("Function trigger failed: " + response.status);
            this.setState({
              apiError: true,
              result: JSON.stringify(response.json),
            })
          }
        })
        .then(data => {
          this.setState({
            apiError: false,
            result: data,
          });
        })
    }
    catch(err) {
      alert("Function trigger failed: " + err);
      return
    }
  }

  render() {
    return (
      <div>
        <h2>Basic info from SSO</h2>
        <p><b>Name:</b> {this.state.userInfo.userName}</p>
        <p><b>E-mail:</b> {this.state.userInfo.preferredUserName}</p>

        {this.state.showLoginBtn && <button onClick={() => this.loginBtnClick()}>Auth and show</button>}
        <button onClick={() => this.getFunction()}>Get Function!</button>
        <input type="text" onChange={(e) => this.nameChanged(e)} />
        <input type="button" value="Post Function" onClick={() => this.postFunction()} />
        <p><b>Function Result:</b> {this.state.result}</p>

        <p>
          <h2>Profile from Microsoft Graph</h2>
          <div>
            <div><b>Name:</b> {this.state.profile.displayName}</div>
            <div><b>Job title:</b> {this.state.profile.jobTitle}</div>
            <div><b>E-mail:</b> {this.state.profile.mail}</div>
            <div><b>UPN:</b> {this.state.profile.userPrincipalName}</div>
            <div><b>Object id:</b> {this.state.profile.id}</div>
          </div>
        </p>

        <p>
          <h2>User Photo from Microsoft Graph</h2>
          <div >
            {this.state.photoObjectURL && <img src={this.state.photoObjectURL} alt="" />}
          </div>
        </p>

      </div>
    );
  }
}
export default Tab;