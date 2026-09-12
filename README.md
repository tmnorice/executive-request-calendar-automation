# Executive Request + Calendar Automation

A Google Apps Script workflow that turns structured request data in Google Sheets into calendar activity and writes the result back to the request tracker.

## Problem

Scheduling requests often require the same repeated steps: review the request, decide whether calendar action is needed, create the event, and update the request record. Repeating those steps manually increases the chance of missed follow-up and inconsistent tracking.

## What the automation does

The workflow reads request rows from a Google Sheet and applies simple routing logic:

1. Read the request record.
2. Check whether calendar action is required.
3. If no calendar action is required, mark the request as complete.
4. If calendar action is required, build the start and end date/time.
5. Create the event in Google Calendar.
6. Write the calendar status and event ID back to the Sheet.
7. Prevent duplicate events by checking for an existing event ID or created status.

## Tools

- Google Sheets
- Google Apps Script
- Google Calendar

## Sheet structure

The script expects these headers in the `Requests` sheet:

- Request ID
- Date Received
- Request Type
- Client / Requestor
- Email
- Requested Date
- Start Time
- End Time
- Notes
- Calendar Required
- Calendar Status
- Calendar Event ID
- Automation Status

## Automation trigger

The working version is configured with a time-driven Apps Script trigger so requests are processed automatically rather than requiring a manual Run action.

## Test cases

The workflow was tested with both routing paths:

- `Calendar Required = Yes` → calendar event created, event ID written back, automation marked complete.
- `Calendar Required = No` → calendar creation skipped, request marked not required and complete.

## Notes

This repository contains a sanitized portfolio version of the workflow. It does not include private account credentials, client data, API keys, or internal business records.
