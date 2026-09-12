const CONFIG = {
  SHEET_NAME: 'Requests',
  TIME_ZONE: 'America/Phoenix'
};

function processRequests() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Requests sheet not found.');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  function columnIndex(headerName) {
    const index = headers.indexOf(headerName);

    if (index === -1) {
      throw new Error(`Missing column: ${headerName}`);
    }

    return index;
  }

  const columns = {
    requestId: columnIndex('Request ID'),
    requestType: columnIndex('Request Type'),
    client: columnIndex('Client / Requestor'),
    email: columnIndex('Email'),
    requestedDate: columnIndex('Requested Date'),
    startTime: columnIndex('Start Time'),
    endTime: columnIndex('End Time'),
    notes: columnIndex('Notes'),
    calendarRequired: columnIndex('Calendar Required'),
    calendarStatus: columnIndex('Calendar Status'),
    eventId: columnIndex('Calendar Event ID'),
    automationStatus: columnIndex('Automation Status')
  };

  for (let rowNumber = 1; rowNumber < data.length; rowNumber++) {
    const row = data[rowNumber];

    const requestId = String(row[columns.requestId] || '').trim();

    if (!requestId) {
      continue;
    }

    const calendarRequired =
      String(row[columns.calendarRequired] || '').trim().toLowerCase();

    const existingEventId =
      String(row[columns.eventId] || '').trim();

    const existingCalendarStatus =
      String(row[columns.calendarStatus] || '').trim().toLowerCase();

    if (existingEventId || existingCalendarStatus === 'created') {
      continue;
    }

    if (calendarRequired !== 'yes') {
      sheet
        .getRange(rowNumber + 1, columns.calendarStatus + 1)
        .setValue('Not Required');

      sheet
        .getRange(rowNumber + 1, columns.automationStatus + 1)
        .setValue('Complete');

      continue;
    }

    try {
      const startDateTime = buildDateTime(
        row[columns.requestedDate],
        row[columns.startTime]
      );

      const endDateTime = buildDateTime(
        row[columns.requestedDate],
        row[columns.endTime]
      );

      if (endDateTime <= startDateTime) {
        throw new Error('End time must be later than start time.');
      }

      const requestType =
        String(row[columns.requestType] || '').trim();

      const client =
        String(row[columns.client] || '').trim();

      const email =
        String(row[columns.email] || '').trim();

      const notes =
        String(row[columns.notes] || '').trim();

      const eventTitle =
        [requestType, client].filter(Boolean).join(' — ') ||
        `Request ${requestId}`;

      const description = [
        `Request ID: ${requestId}`,
        email ? `Requestor Email: ${email}` : '',
        notes ? `Notes: ${notes}` : ''
      ]
        .filter(Boolean)
        .join('\n');

      const calendar = CalendarApp.getDefaultCalendar();

      const event = calendar.createEvent(
        eventTitle,
        startDateTime,
        endDateTime,
        {
          description: description
        }
      );

      sheet
        .getRange(rowNumber + 1, columns.calendarStatus + 1)
        .setValue('Created');

      sheet
        .getRange(rowNumber + 1, columns.eventId + 1)
        .setValue(event.getId());

      sheet
        .getRange(rowNumber + 1, columns.automationStatus + 1)
        .setValue('Complete');

    } catch (error) {
      sheet
        .getRange(rowNumber + 1, columns.calendarStatus + 1)
        .setValue('Error');

      sheet
        .getRange(rowNumber + 1, columns.automationStatus + 1)
        .setValue(error.message);
    }
  }
}

function buildDateTime(dateValue, timeValue) {
  let dateText;
  let timeText;

  if (dateValue instanceof Date) {
    dateText = Utilities.formatDate(
      dateValue,
      CONFIG.TIME_ZONE,
      'yyyy-MM-dd'
    );
  } else {
    dateText = String(dateValue).trim();
  }

  if (timeValue instanceof Date) {
    timeText = Utilities.formatDate(
      timeValue,
      CONFIG.TIME_ZONE,
      'h:mm a'
    );
  } else {
    timeText = String(timeValue).trim();
  }

  return Utilities.parseDate(
    `${dateText} ${timeText}`,
    CONFIG.TIME_ZONE,
    'yyyy-MM-dd h:mm a'
  );
}
