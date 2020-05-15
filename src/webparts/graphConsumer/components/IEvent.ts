export interface IEvent {
  subject: string;
  body: IMailBody;
  start: ITime;
  end: ITime;
  location: ILocation;
  attendees: IAttendee[];
}
export interface IEventListItem {
  id: number;
  subject: string;
  body: IMailBody;
  start: ITime;
  end: ITime;
  location: ILocation;
  requiredattendees: IAttendee[];
  optionalattendees: IAttendee[];
}
export interface IAttendee {
  emailAddress: { address: string; name?: string };
  type: string;
}
export interface ITime {
  dateTime: string;
  timeZone: string;
}
export interface IMailBody {
  contentType: string;
  content: string;
}
export interface ILocation {
  displayName: string;
  locationType?: string;
  uniqueId?: string;
  uniqueIdType?: string;
}
