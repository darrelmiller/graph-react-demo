import * as GraphModels from "@microsoft/microsoft-graph-types"
import React, { useState, useEffect } from 'react';
import { Table } from 'reactstrap';
import moment from 'moment';
import GraphService from './GraphService';
import { Client } from "@microsoft/microsoft-graph-client";

// Helper function to format Graph date/time

interface ICalendarProps {
  client: Client;
  showError: (message: string, debug: string) => void;
}

const Calendar = (props: ICalendarProps) => {
  const formatDateTime = (dateTime: any): string => {
    return moment.utc(dateTime).local().format('M/D/YY h:mm A');
  }
  const [events, setEvents] = useState<GraphModels.Event[]>([]);

  useEffect(() => {
    if(events.length === 0) {
      new GraphService(props.client).getEvents().then((evts) =>
        setEvents(evts)
      );
    }
  }, [events, props.client]);
    return (
      <div>
        <h1>Calendar</h1>
        <Table>
          <thead>
            <tr>
              <th scope="col">Organizer</th>
              <th scope="col">Subject</th>
              <th scope="col">Start</th>
              <th scope="col">End</th>
            </tr>
          </thead>
          <tbody>
            {events.map(
              function(event){
                return(
                  <tr key={event.id}>
                    <td>{event?.organizer?.emailAddress?.name}</td>
                    <td>{event.subject}</td>
                    <td>{formatDateTime(event?.start?.dateTime)}</td>
                    <td>{formatDateTime(event?.end?.dateTime)}</td>
                  </tr>
                );
              })}
          </tbody>
        </Table>
      </div>
    );
  };
  export default Calendar;