import * as GraphModels from "@microsoft/microsoft-graph-types"
import React, { useState, useEffect } from 'react';
import { Table } from 'reactstrap';
import moment from 'moment';
import { getMessages } from './GraphService';
import { Client } from "@microsoft/microsoft-graph-client";
interface IMessagesProps {
  client: Client;
  showError: (message: string, debug: boolean) => void;
}

const Messages = (props: IMessagesProps) => {
  // Helper function to format Graph date/time
  const formatDateTime = (dateTime: any): string => {
    return moment.utc(dateTime).local().format('M/D/YY h:mm A');
  }
  const [messages, setMessages] = useState<GraphModels.Message[]>([]);
  useEffect(() => {
    if(messages.length === 0) {
      getMessages(props.client).then((mess) =>
        setMessages(mess)
      );
    }
  }, [messages, props.client]);
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
            {messages.map(
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
;}
export default Messages;