
//Define calendar(s): addCalendar ("Unique Calendar Name", "Window title", "Form element's name", Form name")
addCalendar("FechaCal", "Calendar", "Fecha", "FrontPage_Form1");
addCalendar("FechaCalU", "Calendar", "Fecha", "FrontPage_Form1");
addCalendar("FechaCalR", "Calendar", "FECHA", "FrontPage_Form1");
addCalendar("FechaCalRU", "Calendar", "FECHA", "FrontPage_Form1");
addCalendar("FechaCalD", "Calendar", "FechaD", "FrontPage_Form1");
addCalendar("FechaCalH", "Calendar", "FechaH", "FrontPage_Form1");

// default settings for English
// Uncomment desired lines and modify its values
 setFont("verdana", 9);
setWidth(150, 1, 15, 1);
setFormat("dd-mon-yyyy");
// setFormat("dd-mon-yyyy");
 setSize(200, 176, -200, 16);

// setWeekDay(0);
 setMonthNames("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
 setDayNames("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
 setLinkNames("[Close]", "[Clear]");
