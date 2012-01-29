using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Calendar;


namespace ConsoleApplication1
{
    class GoogleCalendarApi
    {
        CalendarService service;

        public GoogleCalendarApi()
        {
            AuthenticateV2();
        }
        public void GetUserCalendars(String emailAddress)
        {

        }
        public void AuthenticateV3()
        {
            //The API manual states that I should use an API key and Oauth2LeggedAuthenticator,
            //var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            //provider.ClientIdentifier = ClientCredentials.ClientID;
            //provider.ClientSecret = ClientCredentials.ClientSecret;

            //var authenticator = new OAuth2LeggedAuthenticator(ClientCredentials.ClientID, ClientCredentials.ClientSecret, "myworkusername", "workdomain.com");
            //Google.Apis.Calendar.v3.CalendarService service = new Google.Apis.Calendar.v3.CalendarService(authenticator);
            //service.Key = ClientCredentials.ApiKey;
            //var result = service.CalendarList.List().Fetch();
            //Assert.IsTrue(result.Items.Count > 0);

        }
        public void AuthenticateV2()
        {
            // create an OAuth factory to use
            GOAuthRequestFactory requestFactory = new GOAuthRequestFactory("cl", "MyApp");
            requestFactory.ConsumerKey = ConfigurationManager.AppSettings["consumerKey"];
            requestFactory.ConsumerSecret = ConfigurationManager.AppSettings["consumerSecret"];

            service = new CalendarService("MyApp");
            service.RequestFactory = requestFactory;
            
        }

        public void GetCalendarFeed()
        {
            // example of performing a query (use OAuthUri or query.OAuthRequestorId)
            Uri calendarUri = new OAuthUri("http://www.google.com/calendar/feeds/default/allcalendars/full", "david.turner_adm", ConfigurationManager.AppSettings["consumerKey"]); //allcalendars or owncalendars
      
            // https://www.googleapis.com/calendar/v3/users/userId/calendarList

            CalendarQuery query = new CalendarQuery();
            query.Uri = calendarUri;
            // query.OAuthRequestorId = "audit.admin@mandc-test.com"; // can do this instead of using OAuthUri for queries
            CalendarFeed feed = service.Query(query);
            foreach (CalendarEntry entry in feed.Entries)
            {
                Console.WriteLine(entry.Title.Text);
            }
            Console.WriteLine("Query Success!");
            Console.ReadLine();
        }

        public void WriteEventInformation()
        {
            EventQuery query = new EventQuery();
            query.Uri = new Uri("https://www.google.com/calendar/feeds/david.turner_adm@mandc-test.com/private/full?xoauth_requestor_id=david.turner_adm@mandc-test.com");
            // Tell the service to query:
            EventFeed calFeed = service.Query(query);
            foreach (EventEntry entry in calFeed.Entries)
            {
              
              //  GetCalendarId(entry);
              //  GetEventId(entry);
              //  WriteCreatorDisplayName(entry);
              //  WriteEventId(entry);
              //  WriteUpdatedDateTimeW(entry);
              //  WriteDescription(entry);
              //  WriteTitleOrSummary(entry);
              //  WriteEventStatus( entry);
              //  WriteRecurrenceDetails(entry);
              //  WriteTransparency(entry);
              //  WriteVisibility(entry);
              //  WriteStartDateAndTime(entry);
              //  WriteEndDateAndTime(entry);
              //  WriteEventAllDay(entry);
              //  WriteLocation(entry);
              //  WriteEndTimeZone(entry);
              //  WriteReminders(entry);
              //  WriteAttendees(entry);
                WriteGuestsCanModify(entry);

                Console.WriteLine();
                Console.WriteLine(entry.Title.Text);
            }
        }

        public void GetUserCalendarList()
        {
            //https://www.googleapis.com/calendar/v3/users/userId/calendarList //The Google Calendar API currently only supports the special userId value me, indicating the currently authenticated user.
        }

        private void WriteCreatorDisplayName(EventEntry entry)
        {
            Console.WriteLine("********  DISPLAY NAME    ***********");     
            Console.WriteLine(entry.Authors[0].Name);
            Console.WriteLine("*******************");

        }
        private void WriteCreatorEmail(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            Console.WriteLine(entry.Authors[0].Email);
            Console.WriteLine("*******************");

        }
        private void WriteEventId(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            Console.WriteLine(entry.EventId);
            Console.WriteLine("*******************");

        }
        private void WriteUpdatedDateTimeW(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            Console.WriteLine(entry.Updated);
            Console.WriteLine("*******************");
        }
        /****PROB NOT NEED *****/
        private void WriteEventKind(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            Console.WriteLine("*******************");
        }
        /*** CANT GET ANYTHING MEANINGFUL ******/
        private void WriteDescription(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            Console.WriteLine(entry.Content.Src);
            Console.WriteLine("*******************");
        }

        private void WriteTitleOrSummary(EventEntry entry)
        {
            Console.WriteLine("********TITLE OR SUMMARY***********");
            Console.WriteLine(entry.Title.Text);
            Console.WriteLine("*******************");
        }
        private void WriteEventStatus(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            Console.WriteLine(entry.Status.Value); //i.e. event confirmed
            Console.WriteLine("*******************");
        }
        /***CANT AS YET FIND THIS VALUE****/
        private void WritePrivateValue(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            Console.WriteLine();
            Console.WriteLine("*******************");
        }
        private void WriteRecurrenceDetails(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            if (entry.Recurrence!=null && entry.Recurrence.Value!=null)
            {
                Console.WriteLine(entry.Recurrence.Value);// DTSTART:20120128T090000Z     DTEND:20120128T100000Z     RRULE:FREQ=WEEKLY;INTERVAL=5;BYDAY=MO,SA
            }           
            Console.WriteLine("*******************");
        }
        private void WriteTransparency(EventEntry entry)
        {
            Console.WriteLine("******** TRANSPARENCY    ***********");
            Console.WriteLine(entry.EventTransparency.Value); //i.e. opaque, transparent
            Console.WriteLine("*******************");
        }
        private void WriteVisibility(EventEntry entry)
        {
            Console.WriteLine("******** VISIBILITY       ***********");
            Console.WriteLine(entry.EventVisibility.Value); // private, default or public
            Console.WriteLine("*******************");
        }
        private void WriteStartDateAndTime(EventEntry entry)
        {
            Console.WriteLine("******** DATE EVENT BEGINS ***********");
            if (entry.Times != null && entry.Times.Count > 0)
            {
                Console.WriteLine(entry.Times[0].StartTime);
            }
            
            Console.WriteLine("*******************");
        }
        private void WriteEndDateAndTime(EventEntry entry)
        {
            Console.WriteLine("******** DATE EVENT ENDS   ***********");
            if (entry.Times != null && entry.Times.Count > 0)
            {
                Console.WriteLine(entry.Times[0].EndTime);
            }
            Console.WriteLine("*******************");
        }

        private void WriteEventAllDay(EventEntry entry)
        {
            Console.WriteLine("******** IS EVENT ALL DAY   ***********");
            if (entry.Times != null && entry.Times.Count > 0)
            {
                Console.WriteLine(entry.Times[0].AllDay);
            }
            
            Console.WriteLine("*******************");
        }
        /*********STILL NEED TO IMPLEMENT - CANT SEE TIME ZONE IN ENTRY.TIME**********/
        private void WriteEndTimeZone(EventEntry entry)
        {
            Console.WriteLine("********  TIME ZONE      ***********");           
            if (entry.Times != null && entry.Times.Count > 0)
            {               
               Console.WriteLine(entry.Times[0].EndTime); // output gd
            }
            Console.WriteLine("*******************");
        }
        /*********STILL NEED TO IMPLEMENT - CANT SEE TIME ZONE IN ENTRY.TIME**********/
        private void WriteStartTimeZone(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            Console.WriteLine();
            Console.WriteLine("*******************");
        }

        private void WriteReminders(EventEntry entry)
        {
            Console.WriteLine("********  REMINDERS     ***********");
 
            foreach (Reminder rem in entry.Reminders)
            {
                //FOR SOME REASSON REMINDER ARE INCLUDED TWICE SO ONLY INCLUDE FIRST HALF ON ITERATION
                Console.WriteLine("Abosolute time " + rem.AbsoluteTime);// minutes, hours, days, weeks are option in the drop down; email or popup;type number for the minutes
                Console.WriteLine("Days " + rem.Days);// RETURNS 0
                Console.WriteLine("Hours " + rem.Hours);// RETURNS 0
                Console.WriteLine("Minutes " + rem.Minutes);//this is the only one used NOT DAYS OR HOURS OR ABSOLUTE TIME
                Console.WriteLine("Method " + rem.Method);//alert(called pop-up in calendar) or email
            }
            Console.WriteLine("*******************");
        }
        private void WriteLocation(EventEntry entry)
        {
            Console.WriteLine("******** LOCATION    ***********");

            
            foreach (Where loc in entry.Locations)
            {
                Console.WriteLine(loc.ValueString); //location
            }
            Console.WriteLine("*******************");
        }

        /***** WHAT ARE CONTRIBUTES?******/
        private void WriteAttendees(EventEntry entry)
        {
            Console.WriteLine("********  ATTENDEES      ***********");

            foreach (Who cont in entry.Participants)
            {
                //AttendeeStatus = EVENT_ACCEPTED, EVENT_DECLINED, EVENT_INVITED, EVENT_TENTATIVE
                //AttendeeType = EVENT_OPTIONAL, EVENT_REQUIRED
                //RelType = EVENT_ATTENDEE, EVENT_ORGANIZER, EVENT_PERFORMER, EVENT_SPEAKER
                                                //MESSAGE_BCC, MESSAGE_CC, MESSAGE_FROM, MESSAGE_REPLY_TO, MESSAGE_TO, TASK_ASSIGNED_TO
                if (cont.Attendee_Status != null)
                {
                    Console.WriteLine("Status = " + cont.Attendee_Status.Value);                 
                }
             
                if (cont.Attendee_Type!=null)
                {
                    Console.WriteLine("Type = " + cont.Attendee_Type.Value);//DOESNT RETURN ANY VALUE??? should be optional or required
                }
                
                Console.WriteLine("email = " + cont.Email);
                Console.WriteLine("rel = " + cont.Rel);
                Console.WriteLine("Name of attendee = " + cont.ValueString);
            }
            
            Console.WriteLine("*******************");
        }
        /***STILL NEED TO IMPLEMENT******/
        private void WriteGuestsCanModify(EventEntry entry)
        {
            Console.WriteLine("********   GUESTS CAN MODIFY??     ***********");
                   Console.WriteLine(entry.);         
            Console.WriteLine("*******************");
        }
         /***STILL NEED TO IMPLEMENT******/
        private void WriteGuestsCanInviteOthers(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            Console.WriteLine();
            Console.WriteLine("*******************");
        }
          /***STILL NEED TO IMPLEMENT******/
        private void WriteGuestsCanSeeOtherGuests(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
            Console.WriteLine();
            Console.WriteLine("*******************");
        }

        private void GetEventId(EventEntry entry)
        {
            Console.WriteLine("********        ***********");
          
            Console.WriteLine(entry.EventId);

            Console.WriteLine("*******************");

        }
        private void GetCalendarId(EventEntry entry)
        {
            Console.WriteLine("********        ***********");

            Console.WriteLine(entry.Id.Uri);

            Console.WriteLine("*******************");

        }
        private void WriteComments()
        {
        }
    }
}
