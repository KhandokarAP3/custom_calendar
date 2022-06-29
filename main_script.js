const CustomCalendar = {
    eventsListTitle: "Events",
    calendar: {},
    eventModal: {},
    spEvents: [],
    init: function() {
        CustomCalendar.retrieveEvents().done(function(spEvents) {
            console.log("SP Events", spEvents);

            CustomCalendar.spEvents = spEvents;
            
            if(spEvents.length > 0) {
                spEvents.forEach(spEvent => {
                    const eventToAdd = {
                        id: spEvent.ID,
                        title: spEvent.Title,
                        start: new Date(spEvent.EventDate),
                        end: spEvent.EndDate != "" && spEvent.EndDate != null ? new Date(spEvent.EndDate) : undefined,
                        allDay: spEvent.fAllDayEvent,
                        backgroundColor: CustomCalendar.generateBGColorByCategory(spEvent.Category),
                        // textColor: "",
    
                    };

                    CustomCalendar.calendar.addEvent(eventToAdd);
                });
            }

            CustomCalendar.calendar.render();
        });

        CustomCalendar.bindEvents();
    },
    bindEvents: function() {
        CustomCalendar.eventModal = new bootstrap.Modal(document.getElementById('eventInfoModal'));
        $(".txtDateControl").datepicker();
    },
    retrieveEvents: function() {
        const promise = $.Deferred();

        $.ajax({
            async: true,
            url: `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/GetByTitle('${CustomCalendar.eventsListTitle}')/items`,
            method: "GET",
            headers: {  
                "accept": "application/json;odata=nometadata",
                "content-type": "application/json;odata=nometadata"
            },  
            success: function(data) {
                let results = [];
                
                if(data != null && data.value != null) {
                    results = data.value;
                }
    
                promise.resolve(results);
            },  
            error: function(error) {  
                console.log(JSON.stringify(error));  
                
                promise.reject();
            }
        });
    
        return promise.promise();
    },
    generateBGColorByCategory: function(categoryName) {
        switch(categoryName) {
            case "Meeting": 
                return "red";
            case "Holiday": 
                return "blue";
            case "Birthday": 
                return "green";
            default: 
                return "#c8c8c8";
        }
    }
}

document.addEventListener('DOMContentLoaded', function() {
    const calendarElement = document.getElementsByClassName('divCustomCalendarMain');
    CustomCalendar.calendar = new FullCalendar.Calendar(calendarElement[0], {
        initialView: 'dayGridMonth',
        headerToolbar: {
            left: 'prev,next today',
            center: 'title',
            right: 'dayGridMonth,timeGridWeek,timeGridDay'
        },
        themeSystem: 'bootstrap5',
        events: [],
        eventClick: function(info) {
            if(info != null && info.event != null) {
                const matchedEvents = CustomCalendar.spEvents.filter(evt => evt.ID == info.event.id);
                if(matchedEvents != null && matchedEvents.length > 0) {
                    const matchedEvent = matchedEvents[0];

                    console.log("Selected Event", matchedEvent);

                    $(".txtEventTitle").val(matchedEvent.Title);
                    $(".txtEventLocation").val(matchedEvent.Location);
                    $(".txtEventDesc").html(matchedEvent.Description);
                    $(".slEventCategory").val(matchedEvent.Category).trigger("change");
                    $("#chkAllDay").prop('checked', matchedEvent.fAllDayEvent);

                    $(".txtEventStartDate").datepicker( "setDate", matchedEvent.EventDate );
                    $(".txtEventEndDate").datepicker( "setDate", matchedEvent.EndDate );

                    CustomCalendar.eventModal.show();
                }
            }
        },
        dateClick: function(info) {
            console.log('Clicked on: ' + info.dateStr);
        }
    });
    CustomCalendar.calendar.render();
});

$(document).ready(function() {
    SP.SOD.executeFunc("sp.js", null, function () {
        SP.SOD.executeOrDelayUntilScriptLoaded(function () {
            CustomCalendar.init();
        }, "sp.js");
    });
});