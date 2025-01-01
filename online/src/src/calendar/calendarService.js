//123112313456465

class CalendarService {
  constructor(calendarId) {
    this.calendarId = calendarId;
    var a = 1;
  }

  createEvent(eventData) {
    return Calendar.Events.insert(eventData, this.calendarId);
  }

  deleteEvent(eventId) {
    return Calendar.Events.remove(this.calendarId, eventId);
  }

  getEvent(eventId) {
    return Calendar.Events.get(this.calendarId, eventId);
  }
}