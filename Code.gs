let KEYWORDS = ["[WORK]"]; // TODO: Have user-assigned tags here
let TASK_IDENTIFIER = "[TASK]";
let DAYS_IN_ADVANCE = 7;
let PRIORITY_ORDER = ["low", "medium", "high"];

let SHEETS_URL = 'https://docs.google.com/spreadsheets/d/1c_82kq2MRTjNHmnnOZcVbA4_fepVFuw1EmuU13iz4WU/edit#gid=0';

/**
 * Schedule the tasks to be scheduled every 30 min.
 *  -> Probably can make it less often but the compute is free
 *  -> The algorithm is linear to the number of tasks * number of work events
 * Check when the calendar is edited to see if the user said yes to any of the tasks
 *  -> Happens for any change so runs very often (but again, compute is free and there is no more specific trigger)
 */
function setupTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  if (triggers.length == 2) {
    throw new Error('Triggers are already setup.');
  }
  ScriptApp.newTrigger('scheduleTasks').timeBased().everyMinutes(30).create();
  ScriptApp
      .newTrigger('onChangeTrigger')
      .forUserCalendar(Session.getActiveUser().getEmail())
      .onEventUpdated()
      .create();
}

/**
 * Creates the homepage for the Google Calendar Add-On.
 * Includes buttons/inputs for: 
 *  -> View Tasks Spreadsheet
 *  -> Reset Tasks
 *  -> Schedule Tasks
 *  -> Task Form & Submit Button
 */
function setupHomepage() {
	const cardHeader = CardService.newCardHeader().setTitle("AutoScheduler");

  const resetTasksAction = CardService.newAction().setFunctionName('resetApp');
	const resetTasksButton = CardService.newTextButton().setText('Clear Tasks').setOnClickAction(resetTasksAction);

  const scheduleTasks = CardService.newAction().setFunctionName('scheduleTasks');
	const scheduleTasksButton = CardService.newTextButton().setText('Schedule Tasks').setOnClickAction(scheduleTasks);

	const taskName = CardService.newTextInput()
    .setFieldName("task_name")
    .setTitle("Task Name: ");

  const dueDate = CardService.newDateTimePicker()
    .setFieldName("due_date")
    .setTitle("Due Date: ");

  const estMin = CardService.newTextInput()
    .setFieldName("est_min")
    .setTitle("Estimated Time (in minutes): ");

  const minBlockSize = CardService.newTextInput()
    .setFieldName("min_block_size")
    .setTitle("Minimum Block Size (in minutes): ");

  const priority = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setTitle("Priority: ")
    .setFieldName("priority")
    .addItem("Low", "low", false)
    .addItem("Medium", "medium", true)
    .addItem("High", "high", false)

  const description = CardService.newTextInput()
    .setTitle("Description (optional): ")
    .setFieldName("description")
    .setMultiline(true);

	const submitAction = CardService.newAction()
    .setFunctionName('addTaskEvent');
	const submitButton = CardService.newTextButton().setText('Submit').setOnClickAction(submitAction);

	const openTasksButton = CardService.newTextButton()
		.setText("View All Tasks")
		.setOpenLink(CardService.newOpenLink()
			.setUrl(SHEETS_URL));

	const cardSection = CardService.newCardSection()
    .addWidget(openTasksButton)
    .addWidget(resetTasksButton)
    .addWidget(scheduleTasksButton)
    .addWidget(taskName)
    .addWidget(dueDate)
    .addWidget(estMin)
    .addWidget(minBlockSize)
    .addWidget(priority)
    .addWidget(description)
    .addWidget(submitButton);

	const card = CardService.newCardBuilder()
		.setName("autoscheduler")
		.setHeader(cardHeader)
		.addSection(cardSection)
		.build();

	return card;
}

/** 
 * Returns HTML code for the Web App
 *  -> Added so that I can add tasks on my phone. 
 *  -> Please excuse my lack of CSS
 */
function doGet() {
  return HtmlService
      .createTemplateFromFile('Index')
      .evaluate();
}

/**
 * Add task event when inputted from the add-on homepag.e
 */
function addTaskEvent(e) {
  let d = e.formInput;
  addTask(d['task_name'], new Date(d['due_date']['msSinceEpoch']), d['est_min'], d['priority'], d['min_block_size'], d['description']);
}

/**
 * Adds task to the spreadsheet 'database'. 
 *  -> The webapp calls this function directly. 
 */
function addTask(taskName, dueDate, estimatedMinutes, priority, minBlockSize, description) {
  var ss = SpreadsheetApp.openByUrl(SHEETS_URL);
  var activeTasksSheet = ss.getSheetByName("ACTIVE_TASKS");
  var archivedTasksSheet = ss.getSheetByName("ARCHIVED_TASKS");
  activeTasksSheet.appendRow([taskName, dueDate, estimatedMinutes, priority, minBlockSize, description]);
  archivedTasksSheet.appendRow([taskName, dueDate, estimatedMinutes, priority, minBlockSize, description]);
  // scheduleTasks();
}

/** 
 * Called whenever there is an edit made to the calendar. 
 * Checks if any of the scheduled tasks have a "YES". Since "YES" assumes completed, 
 * it modifies the [WORK] timeblock and lowers the estimated time. 
 */
function onChangeTrigger() {
  let today = new Date();
  let maxDate = new Date();
  maxDate.setDate(maxDate.getDate() + DAYS_IN_ADVANCE);

  let taskEvents = CalendarApp.getDefaultCalendar().getEvents(today, maxDate, {search: TASK_IDENTIFIER});
  taskEvents.forEach(function(event) {
        // if (event.getGuestList(true).length == 0) {
        //   return;
        // }
    let selfGuest = event.getGuestByEmail(Session.getActiveUser().getEmail());
    if (event.getGuestList(true).length == 0 || selfGuest.getGuestStatus() == CalendarApp.GuestStatus.YES) {

      let timePassed = (event.getEndTime().getTime() - event.getStartTime().getTime()) / (1000 * 60);
      lowerTaskEstimatedTime(event.getTitle(), timePassed);

      splitOverlappingWork(event);
    }
  });
  // scheduleTasks();
}

/**
 * Helper that splits the [WORK] block if a completed [TASK] overlaps with it. 
 * 
 */
function splitOverlappingWork(overlappingEvent) {
    let workEvents = CalendarApp.getDefaultCalendar().getEvents(overlappingEvent.getStartTime(), overlappingEvent.getEndTime(), {search: "[WORK]"});
    workEvents.forEach(function(workEvent) {
      console.log("events found are", workEvent.getTitle());

      // make another work event beforehand (if it exists)
      if (workEvent.getStartTime().getTime() < overlappingEvent.getStartTime().getTime()) {
        var beforeWorkEvent = CalendarApp.createEvent(workEvent.getTitle(), workEvent.getStartTime(), overlappingEvent.getStartTime());
      }
      // make another work event after (if it exists)
      if (workEvent.getEndTime().getTime() > overlappingEvent.getEndTime().getTime()) {
        var afterWorkEvent = CalendarApp.createEvent(workEvent.getTitle(), overlappingEvent.getEndTime(), workEvent.getEndTime());
      }

      workEvent.deleteEvent();
    });
}

/** 
 * Helper that lowers the estimated time to complete a [TASK] after a user says YES.
 */
function lowerTaskEstimatedTime(taskName, timePassed) {
  var ss = SpreadsheetApp.openByUrl(SHEETS_URL);  
  var editSheet = ss.getSheetByName("ACTIVE_TASKS"); 
  var lastRowEdit = editSheet.getLastRow();
  taskName = taskName.substring(TASK_IDENTIFIER.length + 1);

  for (var i = 2; i <= lastRowEdit; i++) { 
    if (editSheet.getRange(i, 1).getValue() == taskName) {
      let estTimeCell = editSheet.getRange('C' + i);
      let newEstTime = estTimeCell.getValue() - timePassed
      if (newEstTime <= 0) {
        editSheet.deleteRow(i);
        // TODO: move the row to an archived sheet
      } else {
        estTimeCell.setValue(newEstTime);
      }
    }
  }
}

/**
 * Looking at a range of time, the calendar adds a sorted list of events to the open [WORK] spots. 
 *  -> TODO: make this into more functions and easier to read
 */
function scheduleTasks() {
  // Defines the calendar event date range to search.
  let today = new Date();
  let maxDate = new Date();
  maxDate.setDate(maxDate.getDate() + DAYS_IN_ADVANCE);
  const userProperties = PropertiesService.getUserProperties();

  resetApp();

  let sortedTasks = getSortedTasks();
  KEYWORDS.forEach(function(keyword) {
    let workTimes = CalendarApp.getDefaultCalendar().getEvents(today, maxDate, {search: keyword});
    workTimes.forEach(function(event) {
      var tasksIndex = 0;
      var currentTime = event.getStartTime().getTime() > today.getTime() ? event.getStartTime() : today;
      var taskName = TASK_IDENTIFIER + " " + sortedTasks[tasksIndex][0];

      scheduledTimes = getScheduledTimesWithTask(taskName, userProperties);

      while (tasksIndex < sortedTasks.length) {
        taskName = TASK_IDENTIFIER + " " + sortedTasks[tasksIndex][0];
        var taskTimeMinutes = sortedTasks[tasksIndex][2] - scheduledTimes[taskName];
        var blockSize = sortedTasks[tasksIndex][4];

        minBlockSize = blockSize > taskTimeMinutes ? taskTimeMinutes : blockSize;

        if (minBlockSize > 0 && currentTime.getTime() + minBlockSize * 60000 <= event.getEndTime().getTime()) {
          var remainingWorkTime = (event.getEndTime().getTime() - currentTime.getTime()) / (1000 * 60);
          var blockSize = taskTimeMinutes < remainingWorkTime ? taskTimeMinutes : remainingWorkTime;

          var endTime = new Date(currentTime.getTime() + blockSize * 60000);

          var taskEvent = CalendarApp.getDefaultCalendar().createEvent(taskName, currentTime, endTime, {guests: Session.getActiveUser().getEmail()});
          taskEvent.setDescription(sortedTasks[tasksIndex][5]);

          changeScheduledTaskTime(taskName, blockSize);

          currentTime = new Date(currentTime.getTime() + blockSize * 60000);
        }

        tasksIndex += 1;
        if (tasksIndex < sortedTasks.length) {
          scheduledTimes = getScheduledTimesWithTask(TASK_IDENTIFIER + " " + sortedTasks[tasksIndex][0], userProperties);
          taskTimeMinutes = sortedTasks[tasksIndex][2] - scheduledTimes[sortedTasks[tasksIndex][0]];
        }
      }
    });
  });
}

/**
 * Examines `userProperties` to get time scheduled sofar for a [TASK].
 */
function getScheduledTimesWithTask(currentTaskName, userProperties) {
  var scheduledTimes = userProperties.getProperties();
  if (!(currentTaskName in scheduledTimes)) {
    userProperties.setProperty(currentTaskName, 0);
    scheduledTimes = userProperties.getProperties();
  }
  return scheduledTimes;
}

/** 
 * Clears all [TASK] events from yesterday to `DAYS_IN_ADVANCE` tasks.
 */
function resetApp() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();

  // delete all the calendar events with `TASK_IDENTIFIER` on it
  let today = new Date();
  today.setDate(today.getDate() - 1);
  let maxDate = new Date();
  maxDate.setDate(maxDate.getDate() + DAYS_IN_ADVANCE);
  let taskEvents = CalendarApp.getDefaultCalendar().getEvents(today, maxDate, {search: TASK_IDENTIFIER});
  taskEvents.forEach(function(event) {
    if (event.getGuestList(true).length == 0 || event.getGuestByEmail(Session.getActiveUser().getEmail()).getGuestStatus() == CalendarApp.GuestStatus.YES) {
      return;
    }
    event.deleteEvent();
  })
}

/** 
 * Add/subtracts the newly scheduledTime from `userProperties` for a given task.  
 */
function changeScheduledTaskTime(taskName, scheduledTime) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    var scheduledTimeSofar = parseInt(userProperties.getProperty(taskName));
    userProperties.setProperty(taskName, scheduledTime + scheduledTimeSofar);
  } catch (err) {
    Logger.log('FAILED WITH ERROR %s', err.message);
  }
}

/** 
 * Sorts tasks in the following way: due date (asc.), priority (desc.), and estimated time (asc.)
 */
function getSortedTasks() {
  var data = SpreadsheetApp.openByUrl(SHEETS_URL).getDataRange().getValues();
  data.shift();
  data.sort(function(task1, task2) {
    // due date
    var task1Date = new Date(task1[1]).getTime();
    var task2Date = new Date(task2[1]).getTime();
    if (task1Date == task2Date) {
      // priority
      if (task1[3] == task2[3]) {
        // estimated time
        console.log(task1[2] - task2[2])
        return task1[2] - task2[2];
      }
      return PRIORITY_ORDER.indexOf(task2[3]) - PRIORITY_ORDER.indexOf(task1[3]);
    }
    return task1Date < task2Date ? -1: 1;
  });
  return data;
}

function getMinutesFromTime(time) {
  return time / (60 * 100);
}
