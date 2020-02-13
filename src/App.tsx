import React, { useState, useEffect, useCallback } from "react";
import { BrowserRouter as Router, Route } from "react-router-dom";
import { Container } from "reactstrap";
import NavBar, { IUser } from "./NavBar";
import ErrorMessage, { IErrorMessageProps } from "./ErrorMessage";
import Welcome from "./Welcome";
import "bootstrap/dist/css/bootstrap.css";
import config from "./Config";
import { UserAgentApplication } from "msal";
import GraphService from "./GraphService";
import Calendar from "./Calendar";
import Messages from "./messages";
import { ImplicitMSALAuthenticationProvider } from "@microsoft/microsoft-graph-client/lib/es/ImplicitMSALAuthenticationProvider";
import * as Graph from "@microsoft/microsoft-graph-client";
import Config from "./Config";
const App = () => {
  const userAgentApplication = new UserAgentApplication({
    auth: {
      clientId: config.appId,
      redirectUri: config.redirectUri
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true
    }
  });
  const getClient = (userAgentApplication: UserAgentApplication) => {
    const graphAuthProvider = new ImplicitMSALAuthenticationProvider(
      userAgentApplication,
      {
        scopes: Config.scopes
      }
    );
    const client = Graph.Client.initWithMiddleware({
      authProvider: graphAuthProvider
    });
    return client;
  };

  const client = getClient(userAgentApplication);
  const idUser = userAgentApplication.getAccount();
  const [isAuthenticated, setIsAuthenticated] = useState(idUser !== null);
  const [error, setError] = useState<IErrorMessageProps | undefined>(undefined);
  const [user, setUser] = useState<IUser | undefined>(undefined);
  const setErrorMessage = (message: string, debug: string): void => {
    setError({ message, debug });
  };
  const login = async () => {
    try {
      await userAgentApplication.loginPopup({
        scopes: config.scopes,
        prompt: "select_account"
      });

      await getUserProfile();
    } catch (err) {
      var error: IErrorMessageProps = { message: "", debug: "" };

      if (typeof err === "string") {
        var errParts = err.split("|");
        error =
          errParts.length > 1
            ? { message: errParts[1], debug: errParts[0] }
            : { message: err, debug: "" };
      } else {
        error = {
          message: err.message,
          debug: JSON.stringify(err)
        };
      }
      setIsAuthenticated(false);
      setError(error);
      setUser(undefined);
    }
  };
  
  const getUserProfile = useCallback(async () => {
    try {
      // Get the user's profile from Graph
      var user = await new GraphService(client).getUserDetails();
      setIsAuthenticated(true);
      setError(undefined);
      setUser({
        displayName: user.displayName as string,
        email: (user.mail || user.userPrincipalName) as string,
        avatar: ""
      });
    } catch (err) {
      var error: IErrorMessageProps = { message: "", debug: "" };
      if (typeof err === "string") {
        var errParts = err.split("|");
        error =
          errParts.length > 1
            ? { message: errParts[1], debug: errParts[0] }
            : { message: err, debug: "" };
      } else {
        error = {
          message: err.message,
          debug: JSON.stringify(err)
        };
      }
      setIsAuthenticated(false);
      setError(error);
      setUser(undefined);
    }
  }, [client]);
  useEffect(() => {
    if (idUser) {
      // Enhance user object with data from Graph
      getUserProfile();
    }
  }, [idUser, userAgentApplication, getUserProfile]);

  const logout = () => {
    userAgentApplication.logout();
  };

  let errorComponent = null;
  if (error) {
    errorComponent = (
      <ErrorMessage message={error.message} debug={error.debug} />
    );
  }

  return (
    <Router>
      <div>
        <NavBar
          isAuthenticated={isAuthenticated}
          authButtonMethod={isAuthenticated ? () => logout() : () => login()}
          user={user}
        />
        <Container>
          {errorComponent}
          <Route
            exact
            path="/"
            render={props => (
              <Welcome
                {...props}
                isAuthenticated={isAuthenticated}
                user={user}
                authButtonMethod={() => login()}
              />
            )}
          />
          <Route
            exact
            path="/calendar"
            render={props => (
              <Calendar
                {...props}
                showError={setErrorMessage}
                client={client}
              />
            )}
          />
          <Route
            exact
            path="/messages"
            render={props => (
              <Messages
                {...props}
                showError={setErrorMessage}
                client={client}
              />
            )}
          />
        </Container>
      </div>
    </Router>
  );
};

export default App;
