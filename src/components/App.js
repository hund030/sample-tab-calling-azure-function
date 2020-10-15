// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";
import { HashRouter as Router, Route } from "react-router-dom";

import Privacy from "./Privacy";
import TermsOfUse from "./TermsOfUse";
import Tab from "./Tab";

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
function App() {
  // Check for the Microsoft Teams SDK object.
  if (microsoftTeams) {

    // Set app routings that don't require microsoft Teams
    // SDK functionality.  Show an error if trying to access the
    // Home page.
    if (window.parent === window.self) {
      return (
        <Router>
          <Route exact path="/privacy" component={Privacy} />
          <Route exact path="/termsofuse" component={TermsOfUse} />
          <Route exact path="/tab" component={TeamsHostError} />
        </Router>
      );
    }

    // Initialize the Microsoft Teams SDK
    microsoftTeams.initialize(window);

    // Display the app home page hosted in Teams
    return (
      <Router>
        <Route exact path="/tab" component={Tab} />
      </Router>
    );
  }

  // Error when the Microsoft Teams SDK is not found
  // in the project.
  return (
    <h3>Microsoft Teams SDK not found.</h3>
  );
}

/**
 * This component displays an error message in the
 * case when a page is not being hosted within Teams.
 */
class TeamsHostError extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      apiError: false,
      result: "",
      name: "",
    };
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
        <h3 className="Error">Debug your app within the Teams client.</h3>
        <button onClick={() => this.getFunction()}>Get Function!</button>
        <input type="text" onChange={(e) => this.nameChanged(e)} />
        <input type="button" value="Post Function" onClick={() => this.postFunction()} />
        <p><b>Function Result:</b> {this.state.result}</p>
      </div>
    );
  }
}

export default App;
