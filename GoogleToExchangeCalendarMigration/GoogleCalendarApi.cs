using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Calendar;
using log4net;
using log4net.Config;


namespace ConsoleApplication1
{
    class GoogleCalendarApi
    {
        protected static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        CalendarService service;
        ExchangeWriter writer;

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
            log.Debug("Authenticating to calendar service using 2-legged OAuth");
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
                log.Debug("Calendar title from GetCalendarFeed: " + entry.Title.Text);
            }
        }

        public void WriteEventInformation()
        {
            log.Debug("Attempting to write all calendar information for user.....");
            EventQuery query = new EventQuery();
            query.Uri = new Uri("https://www.google.com/calendar/feeds/david.turner_adm@mandc-test.com/private/full?xoauth_requestor_id=david.turner_adm@mandc-test.com");
            // Tell the service to query:
            EventFeed calFeed = service.Query(query);
            int i = 0;
            writer = new ExchangeWriter();
            List<String> List = new List<String>();
            foreach (EventEntry entry in calFeed.Entries)
            {  
 
                if(List.Contains(entry.EventId) == false){
                List.Add(entry.EventId);
                log.Debug("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%   ");
                log.Debug("Calendar entry number: " + (i++));              
                log.Debug("Calendar entry title = " + entry.Title.Text);
                WriteTitleOrSummary(entry);//DONE - works
                GetCalendarId(entry);//not needed
                GetEventId(entry);//not needed
                WriteCreatorDisplayName(entry);//Cannot directly write it to exchange as exchange treate the organisor as the creator
                WriteEventId(entry);//not needed
                WriteUpdatedDateTimeW(entry);//don't think this is needed 
                WriteDescription(entry); // DONE - works            
                WriteEventStatus( entry); // i.e. confirmed
                WriteRecurrenceDetails(entry);// TODO
                WriteTransparency(entry); //i.e  opaque
                WriteVisibility(entry); //i.e. default
                WriteStartDateAndTime(entry);//DONE - works although need to test dif time zones
                WriteEndDateAndTime(entry);//DONE - works although need to test dif time zones
                WriteEventAllDay(entry);//DONE - works
                WriteLocation(entry); //DONE - works
                WriteEndTimeZone(entry);
                WriteReminders(entry); //DONE - works
                WriteAttendees(entry); //SO FAR HAVE WRITTEN ALL ATTENDEES AS OPTIONAL ATTENDEES
                WriteGuestsCanModify(entry); //STILL NEED TO GET DATA FROM GOOGLE

                writer.SaveAndResetAppointment();
                }
            }
        }

        public void GetUserCalendarList()
        {
            //https://www.googleapis.com/calendar/v3/users/userId/calendarList //The Google Calendar API currently only supports the special userId value me, indicating the currently authenticated user.
        }

        private void WriteCreatorDisplayName(EventEntry entry)
        { 
            log.Debug("Creator display name = " + entry.Authors[0].Name);
        }
        private void WriteCreatorEmail(EventEntry entry)
        {
            log.Debug("Creator email " + entry.Authors[0].Email);
        }
        private void WriteEventId(EventEntry entry)
        {
            log.Debug("Calendar event Id = " + entry.EventId);
        }
        private void WriteUpdatedDateTimeW(EventEntry entry)
        {
            log.Debug("Calendar updated DateTimeW = " + entry.Updated);
        }
        /****PROB NOT NEED *****/
        private void WriteEventKind(EventEntry entry)
        {
            log.Debug("********        ***********");
            log.Debug("*******************");
        }
        /*** CANT GET ANYTHING MEANINGFUL ******/
        private void WriteDescription(EventEntry entry)
        {
            log.Debug("Calendar event description = " + entry.Content.Content);
            writer.WriteDescription(entry.Content.Content);
        }

        private void WriteTitleOrSummary(EventEntry entry)
        {
            log.Debug("Calendar event title or summary = " + entry.Title.Text);
            writer.WriteTitle(entry.Title.Text);

        }
        private void WriteEventStatus(EventEntry entry)
        {
            log.Debug("Calendar event status = " + entry.Status.Value); //i.e. event confirmed
        }
        /***CANT AS YET FIND THIS VALUE****/
        private void WritePrivateValue(EventEntry entry)
        {
            log.Debug("Private value = ");

        }
        private void WriteRecurrenceDetails(EventEntry entry)
        {

            if (entry.Recurrence!=null && entry.Recurrence.Value!=null)             
            {
                log.Debug("Recurrence details = " + entry.Recurrence.Value);// DTSTART:20120128T090000Z     DTEND:20120128T100000Z     RRULE:FREQ=WEEKLY;INTERVAL=5;BYDAY=MO,SA
                writer.WriteRecurrence(entry.Recurrence.Value);
            }           

        }
        private void WriteTransparency(EventEntry entry)
        {
            log.Debug("Transparency (opaque or transparent) = " + entry.EventTransparency.Value); //i.e. opaque, transparent
        }
        private void WriteVisibility(EventEntry entry)
        {
            log.Debug("Calendar event visibility = " + entry.EventVisibility.Value); // private, default or public
        }
        private void WriteStartDateAndTime(EventEntry entry)
        {
            if (entry.Times != null && entry.Times.Count > 0)
            {
                log.Debug("Calendar event start date and time = " + entry.Times[0].StartTime);
                writer.WriteStartDate(entry.Times[0].StartTime);
            }
        }
        private void WriteEndDateAndTime(EventEntry entry)
        {
            if (entry.Times != null && entry.Times.Count > 0)
            {
                log.Debug("Calendar event end date and time = " + entry.Times[0].EndTime);
                writer.WriteEndDate(entry.Times[0].EndTime);
            }
        }

        private void WriteEventAllDay(EventEntry entry)
        {
            if (entry.Times != null && entry.Times.Count > 0)
            {
                log.Debug("Is calendar event an all day event? = " + entry.Times[0].AllDay);
                writer.WriteIsAllDayEvent(entry.Times[0].AllDay);
            }
        }
        /*********STILL NEED TO IMPLEMENT - CANT SEE TIME ZONE IN ENTRY.TIME**********/
        private void WriteEndTimeZone(EventEntry entry)
        {
       
            if (entry.Times != null && entry.Times.Count > 0)
            {               
               log.Debug("End time zone = " + entry.Times[0].EndTime); // output gd
            }
        }
        /*********STILL NEED TO IMPLEMENT - CANT SEE TIME ZONE IN ENTRY.TIME**********/
        private void WriteStartTimeZone(EventEntry entry)
        {
            log.Debug("Start time zone");
        }

        private void WriteReminders(EventEntry entry)
        {
            log.Debug("********  REMINDERS     ***********");
 
            foreach (Reminder rem in entry.Reminders)
            {
                //FOR SOME REASSON REMINDER ARE INCLUDED TWICE SO ONLY INCLUDE FIRST HALF ON ITERATION
                log.Debug("Abosolute time of reminder " + rem.AbsoluteTime);// minutes, hours, days, weeks are option in the drop down; email or popup;type number for the minutes
                log.Debug("Days of reminder " + rem.Days);// RETURNS 0
                log.Debug("Hours of reminder " + rem.Hours);// RETURNS 0
                log.Debug("Minutes of reminder " + rem.Minutes);//this is the only one used NOT DAYS OR HOURS OR ABSOLUTE TIME
                log.Debug("Method of reminder " + rem.Method);//alert(called pop-up in calendar) or email
                writer.WriteReminder(rem.Minutes);
            }

        }
        private void WriteLocation(EventEntry entry)
        {            
            foreach (Where loc in entry.Locations)
            {
                log.Debug("Location = " + loc.ValueString);
                writer.WriteLocation(loc.ValueString);
            }
        }

        /***** WHAT ARE CONTRIBUTES?******/
        private void WriteAttendees(EventEntry entry)
        {
            log.Debug("Writing Attendees");

            foreach (Who cont in entry.Participants)
            {
                //AttendeeStatus = EVENT_ACCEPTED, EVENT_DECLINED, EVENT_INVITED, EVENT_TENTATIVE
                //AttendeeType = EVENT_OPTIONAL, EVENT_REQUIRED
                //RelType = EVENT_ATTENDEE, EVENT_ORGANIZER, EVENT_PERFORMER, EVENT_SPEAKER
                                                //MESSAGE_BCC, MESSAGE_CC, MESSAGE_FROM, MESSAGE_REPLY_TO, MESSAGE_TO, TASK_ASSIGNED_TO
                if (cont.Attendee_Status != null)
                {
                    log.Debug("Attendee status = " + cont.Attendee_Status.Value);
                    if (cont.Attendee_Status.Value.Contains("invited"))
                    {
                    }
                    if (cont.Attendee_Status.Value.Contains("declined"))
                    {
                    }
                    if (cont.Attendee_Status.Value.Contains("tentative"))
                    {
                    }
                    if (cont.Attendee_Status.Value.Contains("accepted"))
                    {

                    }
                }            
                if (cont.Attendee_Type!=null)
                {
                    log.Debug("Attendee type = " + cont.Attendee_Type.Value); //DOESNT RETURN ANY VALUE??? should be optional or required
                }
                log.Debug("Attendee email = " + cont.Email);
                log.Debug("Attendee rel = " + cont.Rel);//attended, organiser
                log.Debug("Name of attendee = " + cont.ValueString);
                writer.WriteOptionalAttendee(cont.Email);
            }
        }
        /***STILL NEED TO IMPLEMENT******/
        private void WriteGuestsCanModify(EventEntry entry)
        {
            log.Debug("Guests can modify? = ");
        }
         /***STILL NEED TO IMPLEMENT******/
        private void WriteGuestsCanInviteOthers(EventEntry entry)
        {
            log.Debug("Guests can invite others? = ");

        }
          /***STILL NEED TO IMPLEMENT******/
        private void WriteGuestsCanSeeOtherGuests(EventEntry entry)
        {
            log.Debug("Guests can see other guests? = ");
        }

        private void GetEventId(EventEntry entry)
        {
            log.Debug("Event Id = " + entry.EventId);
        }
        private void GetCalendarId(EventEntry entry)
        {
            log.Debug("Calendar Id = " + entry.Id.Uri);
        }
        private void WriteComments()
        {
        }
    }
}
