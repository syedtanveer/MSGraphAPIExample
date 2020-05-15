import * as React from "react";
import styles from "./GraphConsumer.module.scss";
import { IGraphConsumerProps } from "./IGraphConsumerProps";
import { MSGraphClient } from "@microsoft/sp-http";
import { IEvent, IAttendee, ITime, IMailBody, ILocation, IEventListItem } from "./IEvent";
import { sp } from "@pnp/sp";
import {
  PrimaryButton
} from 'office-ui-fabric-react';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
} from 'office-ui-fabric-react/lib/DetailsList';
import { loadTheme } from "office-ui-fabric-react/lib/Styling";


//CONSTANTS
const listName: string = "EventList";


//INTERFACES

export interface IGraphConsumerState {
  siteId?: string;
  events: IEventListItem[];
}
export interface IEventDetailListItem {
  key: number;
  id: number;
  subject: string;
  body: string;
  start: string;
  end: string;
  location: string;
  requiredattendees: string;
  optionalattendees: string;
}
export default class GraphConsumer extends React.Component<
  IGraphConsumerProps,
  IGraphConsumerState
  > {
  // Crete an event
  private client: any;
  private _columns: IColumn[];
  private _allItems: IEventDetailListItem[];
  public constructor(props) {
    super(props);
    this.client = null;
    this.state = {
      siteId: "",
      events: [] as IEventListItem[]
    };
    this._allItems =
      [{
        key: 0,
        id: 0,
        subject: "",
        body: "",
        start: "",
        end: "",
        location: "",
        requiredattendees: "",
        optionalattendees: "",
      }];
    this._columns = [{ key: 'column1', name: 'Id', fieldName: 'id', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column2', name: 'Subject', fieldName: 'subject', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column2', name: 'Body', fieldName: 'body', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column2', name: 'Start', fieldName: 'start', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column2', name: 'End', fieldName: 'end', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column2', name: 'Location', fieldName: 'location', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column2', name: 'Required Attendees', fieldName: 'requiredattendees', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column2', name: 'Optional Attndees', fieldName: 'optionalattendees', minWidth: 100, maxWidth: 200, isResizable: true },
    {
      key: 'creteevent', name: 'Create Event', fieldName: 'createevent', minWidth: 100, maxWidth: 200,

      onRender: item => (
        // tslint:disable-next-line:jsx-no-lambda
        <PrimaryButton text="Create Event" data-selection-invoke={true} onClick={() => this.onButtonClickHandler(item.id)} />
      )
    }
    ];
  }

  //LIFE CYCLE METHODS
  public componentDidMount(): void {
    this.props.context.msGraphClientFactory
      .getClient().then((client: MSGraphClient): void => {
        this.client = client;
      }).then(() => {
        sp.web.lists.getByTitle("EventList").items.select("Id", "Title", "Body", "Start", "End", "Location", "Timezone", "RequiredAttendees/EMail", "OptionalAttendees/EMail")
          .expand("RequiredAttendees", "OptionalAttendees")
          .get()
          .then((items) => {
            let eventArr = [] as IEventListItem[];
            items.forEach((item) => {
              let event = {} as IEventListItem;
              event.id = item.Id;
              event.subject = item.Title;

              let mailBody = {} as IMailBody;
              mailBody.content = item.Body;
              mailBody.contentType = "HTML";
              event.body = mailBody;

              let startTime = {} as ITime;
              startTime.dateTime = item.Start;
              startTime.timeZone = item.Timezone;
              event.start = startTime;

              let endTime = {} as ITime;
              endTime.dateTime = item.End;
              endTime.timeZone = item.Timezone;
              event.end = endTime;

              let location = {} as ILocation;
              location.displayName = item.Location;
              event.location = location;

              let reqAttendees = [] as IAttendee[];
              let optAttendees = [] as IAttendee[];
              item.RequiredAttendees.forEach((person) => {
                let attendee = {} as IAttendee;
                attendee.emailAddress = { address: "" }
                attendee.emailAddress.address = person.EMail;
                attendee.type = "required";
                reqAttendees.push(attendee);
              });
              item.OptionalAttendees.forEach((person) => {
                let attendee = {} as IAttendee;
                attendee.emailAddress = { address: "" }
                attendee.emailAddress.address = person.EMail;
                attendee.type = "optional";
                optAttendees.push(attendee);
              });
              event.requiredattendees = reqAttendees;
              event.optionalattendees = optAttendees;
              eventArr.push(event);
            });

            this._allItems = [];
            for (let i = 0; i < eventArr.length; i++) {
              this._allItems = eventArr.map((e) => {
                return {
                  key: e.id,
                  id: e.id,
                  subject: e.subject,
                  body: e.body.content,
                  start: e.start.dateTime,
                  end: e.end.dateTime,
                  location: e.location.displayName,
                  requiredattendees: e.requiredattendees.map((a) => a.emailAddress.address).join(";"),
                  optionalattendees: e.optionalattendees.map((a) => a.emailAddress.address).join(";"),
                };
              });
            }

            this.setState({ events: eventArr });
          });
      });
  }

  private onButtonClickHandler = (id: number) => {
    event.preventDefault();
    let e = this.state.events.filter(e => e.id == id)[0];
    delete e.id;
    e['attendees'] = e.requiredattendees.concat(e.optionalattendees);
    delete e.optionalattendees;
    delete e.requiredattendees;
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client.api('/me/events')
          .post(e).then(res => console.log(res));
      });
  }
  public render(): React.ReactElement<IGraphConsumerProps> {
    return (
      <div className={styles.graphConsumer}>
        <DetailsList
          items={this._allItems}
          columns={this._columns}
          //onRenderItemColumn={this._onRenderItemColumn}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
        />
      </div>
    );
  }
}