const errorHandlerState = {
    eventTitles: [],
    errorMessages: []
};

class ErrorHandler {
    static handleError(error, eventTitle) {
        errorHandlerState.eventTitles.push(eventTitle);
        errorHandlerState.errorMessages.push(error.message);
        
        Logger.log(`Error in event "${eventTitle}": ${error.message}`);
    }

    static overAllError() {
        const errorMessage = `Errors occurred in the following events: ${errorHandlerState.eventTitles.join(', ')} - ${errorHandlerState.errorMessages.join(', ')}`;
        SpreadsheetApp.getActiveSpreadsheet().toast(errorMessage, 'Error');
        Logger.log(errorMessage);
    }
}