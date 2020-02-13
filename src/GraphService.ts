import * as GraphModels from "@microsoft/microsoft-graph-types"
import { PageIterator, PageIteratorCallback, PageCollection } from "@microsoft/microsoft-graph-client";


export async function getUserDetails(client:any) {

  const user = await client.api('/me').get();
  
  return user;
}

export async function getEvents(client:any) {
  
    const events:[GraphModels.Event] = await client
      .api('/me/events')
      .select('subject,organizer,start,end')
      .orderby('createdDateTime DESC')
      .get();
  
    return events;
  }

  export const getMessages = async (client:any): Promise<GraphModels.Message[]> => {
    const messages: PageCollection = await client
      .api('/me/messages')
      .select('subject')
      .orderby('createdDateTime DESC')
      .get();
      console.log(messages);
      return messages.value as GraphModels.Message[];
  }

  export async function getAllMessages(client:any) {
    const messageCollection:PageCollection = await client
      .api('/me/messages')
      .select('subject')
      .orderby('createdDateTime DESC')
      .get();

      let allMessages = new Array<GraphModels.Message>()

      let i = 0;
      let callback: PageIteratorCallback = (data) => {
           allMessages.push(data);
          i++;
          console.log(i);

        return i < 5;
      };
      let pageIterator = new PageIterator(client, messageCollection, callback);
      pageIterator.iterate();
      return { value: allMessages} ;
  }

  export async function sendMail() {

  }