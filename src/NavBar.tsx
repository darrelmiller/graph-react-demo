import React, { useState } from 'react';
import { NavLink as RouterNavLink } from 'react-router-dom';
import {
  Collapse,
  Container,
  Navbar,
  NavbarToggler,
  NavbarBrand,
  Nav,
  NavItem,
  NavLink,
  UncontrolledDropdown,
  DropdownToggle,
  DropdownMenu,
  DropdownItem } from 'reactstrap';
import '@fortawesome/fontawesome-free/css/all.css';

export interface IUser {
  avatar: string;
  displayName: string;
  email: string;
}

interface NavBarProps {
  isAuthenticated: boolean;
  user?: IUser;
  authButtonMethod: () => void;
}
 
const UserAvatar = (props: {user?: IUser}) => {
  // If a user avatar is available, return an img tag with the pic
  if (props.user?.avatar) {
    return <img
            src={props.user.avatar} alt="user"
            className="rounded-circle align-self-center mr-2"
            style={{width: '32px'}}></img>;
  }

  // No avatar available, return a default icon
  return <i
          className="far fa-user-circle fa-lg rounded-circle align-self-center mr-2"
          style={{width: '32px'}}></i>;
}

const AuthNavItem = (props: NavBarProps) => {
  // If authenticated, return a dropdown with the user's info and a
  // sign out button
  if (props.isAuthenticated) {
    return (
      <UncontrolledDropdown>
        <DropdownToggle nav caret>
          <UserAvatar user={props.user}/>
        </DropdownToggle>
        <DropdownMenu right>
          <h5 className="dropdown-item-text mb-0">{props.user?.displayName}</h5>
          <p className="dropdown-item-text text-muted mb-0">{props.user?.email}</p>
          <DropdownItem divider />
          <DropdownItem onClick={props.authButtonMethod}>Sign Out</DropdownItem>
        </DropdownMenu>
      </UncontrolledDropdown>
    );
  }

  // Not authenticated, return a sign in link
  return (
    <NavItem>
      <NavLink onClick={props.authButtonMethod}>Sign In</NavLink>
    </NavItem>
  );
}

export const NavBar = (props: NavBarProps) => {
  const [isOpen, setIsOpen] = useState(false);
  // Only show calendar nav item if logged in
  let calendarLink = null;
  let messagesLink = null;
  if (props.isAuthenticated) {
    calendarLink = (
      <NavItem>
        <RouterNavLink to="/calendar" className="nav-link" exact>Calendar</RouterNavLink>
      </NavItem>
      );
    messagesLink = (
      <NavItem>
        <RouterNavLink to="/messages" className="nav-link" exact>Messages</RouterNavLink>
      </NavItem>
    );
  
  }

  return (
    <div>
      <Navbar color="dark" dark expand="md" fixed="top">
        <Container>
          <NavbarBrand href="/">React Graph Tutorial</NavbarBrand>
          <NavbarToggler onClick={() => setIsOpen(!isOpen)} />
          <Collapse isOpen={isOpen} navbar>
            <Nav className="mr-auto" navbar>
              <NavItem>
                <RouterNavLink to="/" className="nav-link" exact>Home</RouterNavLink>
              </NavItem>
              {calendarLink}
              {messagesLink}
            </Nav>
            <Nav className="justify-content-end" navbar>
              <NavItem>
                <NavLink href="https://developer.microsoft.com/graph/docs/concepts/overview" target="_blank">
                  <i className="fas fa-external-link-alt mr-1"></i>
                  Docs
                </NavLink>
              </NavItem>
              <AuthNavItem
                isAuthenticated={props.isAuthenticated}
                authButtonMethod={props.authButtonMethod}
                user={props.user} />
            </Nav>
          </Collapse>
        </Container>
      </Navbar>
    </div>
  );
};
export default NavBar;