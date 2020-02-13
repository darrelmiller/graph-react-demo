import React, { Component } from 'react';
import { BrowserRouter as Router, Route } from 'react-router-dom';
import { Container } from 'reactstrap';
import NavBar from './NavBar';
import ErrorMessage from './ErrorMessage';
import Welcome from './Welcome';
import 'bootstrap/dist/css/bootstrap.css';
import config from './Config';
import { UserAgentApplication } from 'msal';
import { getUserDetails } from './GraphService.ts';
import Calendar from './Calendar.js';
import Messages from './messages';
import { ImplicitMSALAuthenticationProvider } from "../node_modules/@microsoft/microsoft-graph-client/lib/src/ImplicitMSALAuthenticationProvider";
import * as Graph from "@microsoft/microsoft-graph-client"

class App extends Component {
  constructor(props) {
    super(props);
  
    this.userAgentApplication = new UserAgentApplication({
          auth: {
              clientId: config.appId,
              redirectUri: config.redirectUri
          },
          cache: {
              cacheLocation: "localStorage",
              storeAuthStateInCookie: true
          }
      });
    
    var user = this.userAgentApplication.getAccount();
    this.client = this.getClient(this.userAgentApplication);
      
    this.state = {
      isAuthenticated: (user !== null),
      user: {},
      error: null
    };
  
    if (user) {
      // Enhance user object with data from Graph
      this.getUserProfile();
    }
  }

  async login() {
    try {
      await this.userAgentApplication.loginPopup(
          {
            scopes: config.scopes,
            prompt: "select_account"
        });
       
      await this.getUserProfile();
    }
    catch(err) {
      var error = {};
  
      if (typeof(err) === 'string') {
        var errParts = err.split('|');
        error = errParts.length > 1 ?
          { message: errParts[1], debug: errParts[0] } :
          { message: err };
      } else {
        error = {
          message: err.message,
          debug: JSON.stringify(err)
        };
      }
  
      this.setState({
        isAuthenticated: false,
        user: {},
        error: error
      });
    }
  }

  logout() {
    this.userAgentApplication.logout();
  }

  getClient(userAgentApplication) {
  
    const graphAuthProvider = new ImplicitMSALAuthenticationProvider(userAgentApplication, 
      {
        scopes: ["user.read", "mail.send"]
      });
    const client = Graph.Client.initWithMiddleware({
      authProvider: graphAuthProvider
    });
    return client;
 }

  async getUserProfile() {
    try {
  
        // Get the user's profile from Graph
        var user = await getUserDetails(this.client);
        this.setState({
          isAuthenticated: true,
          user: {
            displayName: user.displayName,
            email: user.mail || user.userPrincipalName
          },
          error: null
        });
      
    }
    catch(err) {
      var error = {};
      if (typeof(err) === 'string') {
        var errParts = err.split('|');
        error = errParts.length > 1 ?
          { message: errParts[1], debug: errParts[0] } :
          { message: err };
      } else {
        error = {
          message: err.message,
          debug: JSON.stringify(err)
        };
      }
  
      this.setState({
        isAuthenticated: false,
        user: {},
        error: error
      });
    }
  }

  render() {
    let error = null;
    if (this.state.error) {
      error = <ErrorMessage message={this.state.error.message} debug={this.state.error.debug} />;
    }

    return (
      <Router>
        <div>
          <NavBar
            isAuthenticated={this.state.isAuthenticated}
            authButtonMethod={this.state.isAuthenticated ? this.logout.bind(this) : this.login.bind(this)}
            user={this.state.user}/>
          <Container>
            {error}
            <Route exact path="/"
              render={(props) =>
                <Welcome {...props}
                  isAuthenticated={this.state.isAuthenticated}
                  user={this.state.user}
                  authButtonMethod={this.login.bind(this)} />
              } />
              <Route exact path="/calendar"
                render={(props) =>
                  <Calendar {...props}
                    showError={this.setErrorMessage.bind(this)}
                    client={this.client} />
                } />
                <Route exact path="/messages"
                render={(props) =>
                  <Messages {...props}
                    showError={this.setErrorMessage.bind(this)}
                    client={this.client} />
                } />
          </Container>
        </div>
      </Router>
    );
  }

  setErrorMessage(message, debug) {
    this.setState({
      error: {message: message, debug: debug}
    });
  }
}

export default App;