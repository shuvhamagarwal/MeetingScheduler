import Router from 'express-promise-router';
import { zonedTimeToUtc } from 'date-fns-tz';
import { findOneIana } from 'windows-iana';
import * as graph from '@microsoft/microsoft-graph-client';
import { DayOfWeek, Event, MailboxSettings } from 'microsoft-graph';
import 'isomorphic-fetch';
import { getTokenOnBehalfOf } from './auth';
import moment from 'moment-timezone'
import { body, validationResult } from 'express-validator';

import validator from 'validator';
import { lastDayOfWeek } from 'date-fns';

//var dict = {};
var seriesMasterId, startDate, endDate;

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

  async function getAccessToken(userId, msalClient) {
    // Look up the user's account in the cache
    try {
      const accounts = await msalClient
        .getTokenCache()
        .getAllAccounts();
  
      const userAccount = accounts.find(a => a.homeAccountId === userId);
  
      // Get the token silently
      const response = await msalClient.acquireTokenSilent({
        scopes: process.env.OAUTH_SCOPES.split(','),
        redirectUri: process.env.OAUTH_REDIRECT_URI,
        account: userAccount
      });
  
      return response.accessToken;
    } catch (err) {
      console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
    }
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


graphRouter.post('/newevent', async function(req, res) {
  const authHeader = req.headers['authorization'];

    if (!authHeader) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {

      
      // Create the event
      try {
        await createEvent(req, res);
      } catch (error) {
        res.status(500).json(error);
      }
      
      try {
        await updateEvent(req, res);
      } catch (error) {
        res.status(500).json(error);
      }

      // Redirect back to the calendar view
      //return res.redirect('/calendarview');
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
  console.log(user_input);
  var temp_dict1:any = {}
  for(i in recurringDays)
  {
    var this_day = recurringDays[i];
    console.log(this_day);
    var this_startTime = user_input[this_day]['startTime'];
    var this_endTime = user_input[this_day]['endTime'];
    if(this_startTime != max_startTime || this_endTime != max_endTime )
    {
      var temp_dict:any = {}
      temp_dict["startTime"] = this_startTime;
      temp_dict["endTime"] = this_endTime;
      temp_dict1[this_day] = temp_dict;
    }
  }
  return temp_dict1
}

async function createEvent(req, res) {
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
        //dict = getToBeUpdatedData(user_input, recurringDays2, max_startTime, max_endTime);
      }
      var interval = 1;
      startDate = req.body['eventStart'];
      endDate = req.body['eventEnd'];
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
    req.body['attendees'].split(';').forEach((attendee: any) => {
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
        .post(newEvent)
        .then(res => {
          seriesMasterId = res.id;
          console.log("SeriesMasterId of this recurrence meeting :" + seriesMasterId);
        });

     
    } catch (error) {
      console.log(error);
      res.status(500).json(error);
    }
  } else {
    // No auth header
    res.status(401).end();
  }
}

async function updateEvent(req, res) {
  const authHeader = req.headers['authorization'];

  if (authHeader) {
    try {
      const client = await getAuthenticatedClient(authHeader);
      const timeZones = await getTimeZones(client);
      const user_input = req.body['formData']
      const isSameTime = req.body['isSameTime']
      var recurringDays2 =  Object.keys(user_input);
      var max_startEndTimes =  mostFrequentTime(user_input);
      var max_startTime = max_startEndTimes[0];
      var max_endTime = max_startEndTimes[1];
      var dict = {}
      if(!isSameTime)
      {
        dict = getToBeUpdatedData(user_input, recurringDays2, max_startTime, max_endTime);
      }
      console.log(dict);
      if(Object.keys(dict).length == 0)
      {
        console.log("All events start on same time. So no update events");
        res.status(201).end();
        return;
      }
      var startDateTime = startDate + "T00:00:00";
      var endDateTime= endDate + "T23:59:59";
      var updatedIds = []
      var isLooping = false;

      console.log("Listing instances between "+ startDateTime + " and " + endDateTime);
      while(true && !isLooping) {
        let instances = await client.api('/me/events/' + seriesMasterId + '/instances?startDateTime='+ startDateTime + '&endDateTime=' + endDateTime).get();
      
        if(instances.value.length == 0){
          console.log("No instances in this timeframe");
          break;
        }

        console.log("Retrieved instances count : " + instances.value.length);

        var i: number;
        for(i = 0; i<instances.value.length; i++)
        {
          var id = instances.value[i].id;
          if(!updatedIds.includes(id))
          { 
            updatedIds.push(id);
          }
          else
          {
            if(i == instances.value.length-1)
            {
              isLooping = true;
              break;
            }
            else
            {
              continue;
            }
          }
          updatedIds.push(id);
          var subject = instances.value[i].subject;
          var oldStartDate = moment(instances.value[i].start.dateTime).format('YYYY-MM-DD');
          //This variable will be used for getting list instances in a  loop
          startDateTime = oldStartDate + "T00:00:00";
          //var oldEndDate = moment(instances.value[i].end.dateTime).local().format('YYYY-MM-DD');
          var day = moment(oldStartDate).format('dddd');
          console.log(day, oldStartDate);
          //console.log( day);
          if(!dict.hasOwnProperty(day))
          {
            console.log("Skipping for " + day);
            continue;
          }
        
          var duration = moment.duration(moment(dict[day].endTime, "HH:mm:ss").diff(moment(dict[day].startTime, "HH:mm:ss")));
          var newStartTime = moment(oldStartDate + " " + dict[day].startTime).format();
          var newEndTime = moment(newStartTime).add(duration);

          console.log("Updated " ,  subject, oldStartDate, day, newStartTime, newEndTime, dict[day].endTime, dict[day].startTime);

          //Build a Graph event
          const updateEvent = {
            start: {
              dateTime: newStartTime,
              timeZone: timeZones.graph
            },
            end: {
              dateTime: newEndTime,
              timeZone: timeZones.graph
            }
          };
      
        // POST /me/events
          await client
            .api('/me/events/'+ id)
            .update(updateEvent)
            .then(res => {
              console.log("Updated the event");
            });
        }
        await delay(500);
     }
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

function delay(ms: number) {
  return new Promise( resolve => setTimeout(resolve, ms) );
}

export default graphRouter;