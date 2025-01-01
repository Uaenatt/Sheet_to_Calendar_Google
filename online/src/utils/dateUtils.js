
class DateUtils {
  static parseDateRangeWithYear(dateRange, baseDate) {
    try {
      // Ensure baseDate is a valid Date object
      if (!(baseDate instanceof Date)) {
        throw new Error("Invalid base date. Must be a Date object.");
      }

      // Normalize input by replacing multiple spaces/tabs with a single space
      dateRange = dateRange.replace(/\s+/g, ' ').trim();
      // Split the input into time range
      var parts = dateRange.split(' '); // Example: "18:00-22:00"
      var timeRange = parts[1]; // The time range part
      var times = timeRange.split('-'); // Split into start and end times

      var startTime = times[0]; // Start time, e.g., "18:00"
      var endTime = times[1]; // End time, e.g., "22:00"

      // Create new Date objects for start and end times using baseDate
      var start = new Date(baseDate);
      var end = new Date(baseDate);

      // Set the time part
      var startParts = startTime.split(':');
      var endParts = endTime.split(':');

      start.setHours(parseInt(startParts[0]), parseInt(startParts[1]), 0, 0); // Set hours and minutes for start
      end.setHours(parseInt(endParts[0]), parseInt(endParts[1]), 0, 0); // Set hours and minutes for end

      // Format the results
      var formatDate = (date) => {
        return date.getFullYear() + '/' +
              ('0' + (date.getMonth() + 1)).slice(-2) + '/' +
              ('0' + date.getDate()).slice(-2) + ' ' +
              ('0' + date.getHours()).slice(-2) + ':' +
              ('0' + date.getMinutes()).slice(-2);
      };

      start = formatDate(start);
      end = formatDate(end);
      Logger.log(`Parsed date range: ${start} - ${end}`);

      if(start.includes('NaN') || end.includes('NaN') ){
        Logger.log("start: " + start + " end: " + end);
        throw new Error("Invalid date range. Unable to parse date and time.");
      }

      return { start, end };

    } catch(e){
      return null;
    }
  }

  static formatDate(date) {
    return date.getFullYear() + '/' +
           ('0' + (date.getMonth() + 1)).slice(-2) + '/' +
           ('0' + date.getDate()).slice(-2) + ' ' +
           ('0' + date.getHours()).slice(-2) + ':' +
           ('0' + date.getMinutes()).slice(-2);
  }
}
