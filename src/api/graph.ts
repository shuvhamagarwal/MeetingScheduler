import Router from 'express-promise-router';
import { zonedTimeToUtc } from 'date-fns-tz';
import { findOneIana } from 'windows-iana';
import * as graph from '@microsoft/microsoft-graph-client';
import { DayOfWeek, Event, MailboxSettings } from 'microsoft-graph';
import 'isomorphic-fetch';
import { getTokenOnBehalfOf } from './auth';
import moment from 'moment-timezone'
import { lastDayOfWeek } from 'date-fns';

async function getAuthenticatedClient(authHeader: string): Promise<graph.Client> {
    const accessToken = await getTokenOnBehalfOf(authHeader);
  
    return graph.Client.init({
      authProvider: (done) => {
        // Call the callback with the
        // access token
        done(null, accessToken!);
      }
    });
  }

  interface TimeZones {
    // The string returned by Microsoft Graph
    // Could be Windows name or IANA identifier.
    graph: string;
    // The IANA identifier
    iana: string;
  }
  
  async function getTimeZones(client: graph.Client): Promise<TimeZones> {
    // Get mailbox settings to determine user's
    // time zone
    const settings: MailboxSettings = await client
      .api('/me/mailboxsettings')
      .get();
  
    // Time zone from Graph can be in IANA format or a
    // Windows time zone name. If Windows, convert to IANA
    const ianaTz = findOneIana(settings.timeZone!);
  
    const returnValue: TimeZones = {
      graph: settings.timeZone!,
      iana: ianaTz ?? settings.timeZone!
    };
  
    return returnValue;
  }
const graphRouter = Router();
graphRouter.get('/calendarview',
  async function(req, res) {
    const authHeader = req.headers['authorization'];

    if (authHeader) {
      try {
        const client = await getAuthenticatedClient(authHeader);

        const viewStart = req.query['viewStart']?.toString();
        const viewEnd = req.query['viewEnd']?.toString();

        const timeZones = await getTimeZones(client);

        // Convert the start and end times into UTC from the user's time zone
        const utcViewStart = zonedTimeToUtc(viewStart!, timeZones.iana);
        const utcViewEnd = zonedTimeToUtc(viewEnd!, timeZones.iana);

        // GET events in the specified window of time
        const eventPage: graph.PageCollection = await client
          .api('/me/calendarview')
          // Header causes start and end times to be converted into
          // the requested time zone
          .header('Prefer', `outlook.timezone="${timeZones.graph}"`)
          // Specify the start and end of the calendar view
          .query({
            startDateTime: utcViewStart.toISOString(),
            endDateTime: utcViewEnd.toISOString()
          })
          // Only request the fields used by the app
          .select('subject,start,end,organizer')
          // Sort the results by the start time
          .orderby('start/dateTime')
          // Limit to at most 25 results in a single request
          .top(25)
          .get();

        const events: any[] = [];

        // Set up a PageIterator to process the events in the result
        // and request subsequent "pages" if there are more than 25
        // on the server
        const callback: graph.PageIteratorCallback = (event) => {
          // Add each event into the array
          events.push(event);
          return true;
        };

        const iterator = new graph.PageIterator(client, eventPage, callback, {
          headers: {
            'Prefer': `outlook.timezone="${timeZones.graph}"`
          }
        });
        await iterator.iterate();

        // Return the array of events
        res.status(200).json(events);
      } catch (error) {
        console.log(error);
        res.status(500).json(error);
      }
    } else {
      // No auth header
      res.status(401).end();
    }
  }
);


graphRouter.post('/newevent',
  async function(req, res) {
    const authHeader = req.headers['authorization'];

    if (authHeader) {
      try {
        const client = await getAuthenticatedClient(authHeader);

        const timeZones = await getTimeZones(client);
        const user_input = req.body['formData']
        const isSameTime = req.body['isSameTime']
        console.log(req.body);
        var recurringDays:DayOfWeek[] =  GetDaysOfWeek(Object.keys(user_input));
        console.log(recurringDays);
        var recurringDays2 =  Object.keys(user_input);
        var max_startEndTimes =  mostFrequentTime(user_input);
        var max_startTime = max_startEndTimes[0];
        var max_endTime = max_startEndTimes[1];
        var duration = moment.duration(moment(max_endTime, "HH:mm:ss").diff(moment(max_startTime, "HH:mm:ss")));
        if(isSameTime)
        {
          var dict = getToBeUpdatedData(user_input, recurringDays2, max_startTime, max_endTime);
        }
        var interval = 1;
        var startDate = req.body['eventStart'];
        var endDate = req.body['eventEnd'];
        // Create a new Graph Event object
        
        
        const newEvent: Event = {
          subject: req.body['eventSubject'],
          start: {
            dateTime: startDate + "T" + max_startTime,
            timeZone: timeZones.graph
          },
          end: {
            dateTime: moment(startDate + "T" + max_startTime).add(duration).format(),
            timeZone: timeZones.graph
          },
          recurrence: {
            pattern: {
              type: 'weekly',
              interval: interval,
              daysOfWeek: recurringDays
            },
            range: {
              type: 'endDate',
              startDate: startDate,
              endDate: endDate
            }
          },
        isOnlineMeeting: true,
        onlineMeetingProvider: 'teamsForBusiness',
        };
        console.log(newEvent);
        // Add attendees if present
    if (req.body['attendees']) {
      newEvent.attendees = [];
      req.body['attendees'].forEach(attendee => {
        newEvent.attendees.push({
          type: 'required',
          emailAddress: {
            address: attendee
          }
        });
      });
    }

        // POST /me/events
        await client.api('/me/events')
          .post(newEvent);

        // Send a 201 Created
        res.status(201).end();
      } catch (error) {
        console.log(error);
        res.status(500).json(error);
      }
    } else {
      // No auth header
      res.status(401).end();
    }
  }
);
// TODO: Implement this router

function GetDaysOfWeek(daysInString:String[]) {
  var x:DayOfWeek[] = [];
  for (let index = 0; index < daysInString.length; index++) {
    const element = daysInString[index];
    switch (element) {
      case "Monday":
        x.push("monday");
        break;
      case "Tuesday":
          x.push("tuesday");
          break;
          case "Wednesday":
            x.push("wednesday");
            break;
          case "Thursday":
              x.push("thursday");
              break;
          case "Friday":
          x.push("friday");
          break;
          case "Saturday":
            x.push("saturday");
            break;
          case "Sunday":
              x.push("sunday");
              break;
    }
  }
  return x;
}

function mostFrequentTime(user_input:any){
  var time_frequency:any = {};
  // Can move to a function, we need multiple return values
  let timeValues:any = Object.values(user_input);
  var i;
  var max_frequency = 0;
  var max_startTime = timeValues[0].startTime;
  var max_endTime= timeValues[0].endTime;
  for(i=0; i<timeValues.length; i++)
  {
    var key = timeValues[i].startTime+timeValues[i].endTime;
    if(time_frequency.hasOwnProperty(key))
    {
      var new_frequency = time_frequency[key]+1;
      time_frequency[key] = new_frequency;
      if(max_frequency<new_frequency)
      {
        max_frequency = new_frequency;
        max_startTime = timeValues[i].startTime;
        max_endTime = timeValues[i].endTime;
      }
    }
    else
    {
      time_frequency[key] = 1;
    }
  }
  return [max_startTime, max_endTime]
}

function getToBeUpdatedData(user_input:any, recurringDays:any, max_startTime:any, max_endTime:any)
{
  var i;
  // After calculating the most frequent time, here we will create an array for those with different times
  for(i in recurringDays)
  {
    var this_day = recurringDays[i];
    var this_startTime = user_input[this_day].startTime;
    var this_endTime = user_input[this_day].endTime;
    if(this_startTime != max_startTime || this_endTime != max_endTime )
    {
      var temp_dict:any = {}
      temp_dict["startTime"] = this_startTime;
      temp_dict["endTime"] = this_endTime;
      temp_dict[this_day] = temp_dict;
    }
  }
  return temp_dict
}


export default graphRouter;