using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
using log4net;
using System.Text.RegularExpressions;



namespace ConsoleApplication1
{
    class ExchangeWriter
    {
        protected static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        ExchangeService service;
        Appointment appt;

        public ExchangeWriter()
        {
            log.Debug("Creating Exchange writer");
            ExchangeConnection con = new ExchangeConnection();
            service = con.GetExchangeService();
            ResetAppointmentObject();
        }

        public void ResetAppointmentObject()
        {
            appt = new Appointment(service);
        }

        public void WriteTitle(String title)
        {
            appt.Subject = title;
        }

        public void SaveAndResetAppointment()
        {
            appt.Save();
            ResetAppointmentObject();
            log.Debug("Appointment has been successfully written to exchange");
        }

        public void WriteOrganiser(String organizer)
        {
          //  appt.Organizer = organizer;
        }

        public void WriteStartDate(DateTime date)
        {
            appt.Start = date;
        }
        public void WriteEndDate(DateTime date)
        {
            appt.End = date;
        }

        public void WriteReminder(int minutesBeforeStart)
        {
            appt.ReminderMinutesBeforeStart = minutesBeforeStart;

        }

        public void WriteIsAllDayEvent(Boolean isAllDay)
        {
            appt.IsAllDayEvent = isAllDay;

        }

        public void WriteLocation(String location)
        {
            appt.Location = location;
        }
        public void WriteDescription(String description)
        {
            appt.Body = description;
        }

        public void WriteOptionalAttendee(String attendee)
        {
            appt.OptionalAttendees.Add(attendee);
        }
        public void WriteRequiredAttendee(String attendee)
        {
            appt.RequiredAttendees.Add(attendee);
        }

        public void WriteRecurrence(String recurrence)
        {
            log.Debug("Writing recurrence to exchange");// DTSTART:20120128T090000Z     DTEND:20120128T100000Z     RRULE:FREQ=WEEKLY;INTERVAL=5;BYDAY=MO,SA

                char[] div = {':','\n','=',';',','};               
                String[] list = recurrence.Split(div); //start, end, rules                
                for(int i = 0; i < list.Length; i++)
                {
                    String entry = list[i];
                    switch (entry)
                    {
                        case "DTSTART":
                            log.Debug("DTSTART =" + list[++i]);
                            break;
                        case "DTEND":
                            log.Debug("DTEND = " + list[++i]);
                            break;
                        case "FREQ":
                            log.Debug("FREQ = " + list[i+1]);
                            switch (list[++i])
                            {
                                //case "DAILY":
                                //    appt.Recurrence = new Recurrence.DailyPattern(appt.Start.Date, 1);
                                //    break;
                                //case "WEEKLY":
                                //    appt.Recurrence = new Recurrence.WeeklyPattern(appt.Start.Date, 1);
                                //    break;
                                //case "MONTHLY":
                                //    appt.Recurrence = new Recurrence.MonthlyPattern(appt.Start.Date, 1);
                                //    break;
                                //case "YEARLY":
                                //    appt.Recurrence = new Recurrence.YearlyPattern(appt.Start.Date, 1);
                                //    break;
                            }
                            
                            break;
                        case "UNTIL":
                            log.Debug("UNTIL= " + list[++i]);
                            break;
                        case "COUNT"://this is number of occurrences of the recurring appointment
                            log.Debug("COUNT = " + list[++i]);
                            break;
                        case "INTERVAL"://this is how often the appointment repeats
                            log.Debug("INTERVAL = " + list[++i]);
                            break;
                        case "BYDAY":
                            log.Debug("BYDAY IS INVOKED");
                            log.Debug("List i + 1 = " + list[i + 1]);
                            if (Regex.IsMatch(list[i+1], @"\d"))
                            {
                                log.Debug("Recurs monthly by day of week: " + list[++i]);
                            }
                            //         appt.Recurrence = new Recurrence.DailyPattern(appt.Start.Date, 1); BYMONTHDAY
                            break;
                        case "BYMONTHDAY":
                            log.Debug("Recurs monthly on:  " + list[++i]);
                            break;
                        case "YEARLY":
                            log.Debug("Recurs yearly");
                  //          appt.Recurrence = new Recurrence.YearlyPattern(appt.Start.Date); 
                            break;
                        case "MO":
                            log.Debug("Recurs on monday");
                            break;
                        case "TU":
                            log.Debug("Recurs on tuesday");
                            break;
                        case "WE":
                            log.Debug("Recurs on wednesday");
                            break;
                        case "TH":
                            log.Debug("Recurs on thursday");
                            break;
                        case "FR":
                            log.Debug("Recurs on friday");
                            break;
                        case "SA":
                            log.Debug("Recurs on saturday");
                            break;
                        case "SU":
                            log.Debug("Recurs on sunday");
                            break;
                    }                    
                }
        }

        public void WriteEndTimeZone(DateTime date)
        {

        }


    }
}
