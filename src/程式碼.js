function main(input, _calendar_id) {
  try {
    const eventService = new EventService(_calendar_id);
    
    switch(input) {
      case "addEvent":
        eventService.addEvent();
        break;
      case "addCustomMenu":
        MenuService.addCustomMenu(_calendar_id);
        break;
      case "logSheetContent":
        SheetUtils.logSheetContent();
        break;
      case "firstUsed":
        SetupUtils.firstUsed();
        break;
      default:
        throw new Error("Function not found");
    }
  } catch(error) {
    ErrorHandler.handleError(error);
    throw error;
  }
}

