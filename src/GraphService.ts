import * as GraphModels from "@microsoft/microsoft-graph-types";
import {
  PageIterator,
  PageIteratorCallback,
  PageCollection,
  Client
} from "@microsoft/microsoft-graph-client";

export default class GraphService {
  constructor(private _client: Client) {}
  public getUserDetails = async (): Promise<GraphModels.User> => {
    const user = await this._client.api("/me").get();

    return user;
  };

  public getEvents = async (): Promise<GraphModels.Event[]> => {
    const events: PageCollection = await this._client
      .api("/me/events")
      .select("subject,organizer,start,end")
      .orderby("createdDateTime DESC")
      .get();

    return events.value as GraphModels.Event[];
  };

  public getMessages = async (): Promise<GraphModels.Message[]> => {
    const messages: PageCollection = await this._client
      .api("/me/messages")
      .select("subject")
      .orderby("createdDateTime DESC")
      .get();
    console.log(messages);
    return messages.value as GraphModels.Message[];
  };

  public getAllMessages = async (): Promise<GraphModels.Message[]> => {
    const messageCollection: PageCollection = await this._client
      .api("/me/messages")
      .select("subject")
      .orderby("createdDateTime DESC")
      .get();

    let allMessages = new Array<GraphModels.Message>();

    let i = 0;
    let callback: PageIteratorCallback = data => {
      allMessages.push(data);
      i++;
      console.log(i);

      return i < 50;
    };
    let pageIterator = new PageIterator(this._client, messageCollection, callback);
    pageIterator.iterate();
    return allMessages;
  };
}
