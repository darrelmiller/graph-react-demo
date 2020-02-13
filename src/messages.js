import React from 'react';
import { Table } from 'reactstrap';
import moment from 'moment';
import { getMessages2 } from './GraphService';

// Helper function to format Graph date/time
function formatDateTime(dateTime) {
  return moment.utc(dateTime).local().format('M/D/YY h:mm A');
}

export default class Messages extends React.Component {
  constructor(props) {
    super(props);

    this.state = {
      messages: []
    };
  }

  async componentDidMount() {
    try {
      console.log("Messages Loaded");
      // Get the user's events
      var messages = await getMessages2(this.props.client);
      // Update the array of events in state
      this.setState({messages: messages});
    }
    catch(err) {
      this.props.showError('ERROR', JSON.stringify(err));
    }
  }

  render() {
    return (
      <div>
        <h1>Messages</h1>
        <Table>
          <thead>
            <tr>
              <th scope="col">Sent</th>
              <th scope="col">Subject</th>
            </tr>
          </thead>
          <tbody>
            {this.state.messages.map(
              (message) => {
                return(
                  <tr key={message.id}>
                    <td>{formatDateTime(message.sentDateTime)}</td>
                    <td>{message.subject}</td>
                  </tr>
                );
              })}
          </tbody>
        </Table>
      </div>
    );
  }
}