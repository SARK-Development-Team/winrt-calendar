using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Windows.ApplicationModel.Appointments;

namespace CalendarFunctions
{
    public class Calendar
    {
        /// <summary>
        /// Creates an Appointment based on the input fields and validates it.
        /// </summary>
        static Appointment CreateAppointment(DateTime date, string EventName, string EventDescription, string StaffName = null, string StaffEmail = null)
        {
            var appointment = new Appointment();

            // StartTime
            var timeZoneOffset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);
            var startTime = new DateTimeOffset(date.Year, date.Month, date.Day, date.Hour,
                date.Minute, 0, timeZoneOffset);
            appointment.StartTime = startTime;

            // Subject
            appointment.Subject = EventName;

            // Details
            appointment.Details = EventDescription;

            // 1 hour duration is selected
            appointment.Duration = TimeSpan.FromHours(2);

            // Reminder
            appointment.Reminder = TimeSpan.FromMinutes(15);

            if (!(StaffName == null))
            {
                appointment.Invitees.Add(new AppointmentInvitee()
                {
                    Address = StaffEmail,
                    Response = AppointmentParticipantResponse.None,
                    DisplayName = StaffName ?? "SARK Insurance Services"
                });
            }

            return appointment;
        }

        /// <summary>
        /// Adds an appointment to the Windows runtime appointment provider.
        /// </summary>
        public static async Task AddEvent(DateTime date, string EventName, string EventDescription, string StaffName = null, string StaffEmail = null)
        {
            var appointment = CreateAppointment(date, EventName, EventDescription, StaffName, StaffEmail);

            // 1. get access to appointmentstore 
            var appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AppCalendarsReadWrite);

            // 2. get calendar 
            AppointmentCalendar sarkcalendar;
            try
            {
                sarkcalendar = await appointmentStore.GetAppointmentCalendarAsync("SARK Move");
            }
            catch (Exception)
            {
                sarkcalendar = await appointmentStore.CreateAppointmentCalendarAsync("SARK Move");
            }

            //  3. add appointment 
            await sarkcalendar.SaveAppointmentAsync(appointment);
        }

        public static async Task<bool> DeleteCalendar(string calendarID)
        {
            var appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadWrite);
            var xml = new XElement("Calendars");

            try
            {
                var calendars = await appointmentStore.FindAppointmentCalendarsAsync();

                foreach (var calendar in calendars)
                {
                    xml.Add(SerializeCalendar(calendar));

                    if (calendar.DisplayName == calendarID)
                    {
                        await calendar.DeleteAsync();
                        return true;
                    }
                }

                xml.Save(@"C:\Users\ramsi\Documents\files.xml");
                return false;

            }
            catch (Exception exc)
            {
                xml.Save(@"C:\Users\ramsi\Documents\files.xml");
                return false;
            }
        }

        static XElement SerializeCalendar(AppointmentCalendar calendar)
        {
            var ele = new XElement("Calendar");
            foreach (var property in typeof(AppointmentCalendar).GetProperties())
            {
                ele.Add(new XElement(property.Name, property.GetValue(calendar)));
            }
            return ele;
        }
    }
}
