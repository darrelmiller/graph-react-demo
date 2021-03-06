import React from 'react';
import {
  Button,
  Jumbotron } from 'reactstrap';
import { IUser } from './NavBar';

const WelcomeContent = (props: IWelcomeProps) => {
  // If authenticated, greet the user
  if (props.isAuthenticated) {
    return (
      <div>
        <h4>Welcome {props.user?.displayName}!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
      </div>
    );
  }

  // Not authenticated, present a sign in button
  return <Button color="primary" onClick={props.authButtonMethod}>Click here to sign in</Button>;
}
interface IWelcomeProps {
  isAuthenticated: boolean;
  authButtonMethod: () => void;
  user?: IUser
}
const Welcome = (props: IWelcomeProps) => {
    return (
      <Jumbotron>
        <h1>React Graph Tutorial</h1>
        <p className="lead">
            This sample app shows how to use the Microsoft Graph API to access Outlook and OneDrive data from React
        </p>
        <WelcomeContent
          isAuthenticated={props.isAuthenticated}
          user={props.user}
          authButtonMethod={props.authButtonMethod} />
      </Jumbotron>
    );
  };
export default Welcome;