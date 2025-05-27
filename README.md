# Calendar Event Extremes Exporter (AppleScript)

This AppleScript exports the earliest and latest events from each of your Apple Calendars to a timestamped CSV file.  
It is designed for general use and produces output that is easy to filter or analyze in spreadsheet applications.

## Features

- **Per-Calendar Analysis:** Reports the earliest and latest events for every calendar in your Apple Calendar.
- **Account-Aware:** Includes the email address associated with each calendar (when available).
- **Human & Machine Friendly:** Outputs a properly escaped CSV with ISO 8601 timestamps.
- **Automatic File Naming:** Saves to your `Documents` folder with a unique timestamped filename.
- **Safe:** Read-only—does not modify, create, or delete any calendar or event data.

## Output

Each run creates a CSV file in your `Documents` folder named like:

```
calendar_event_extremes_YYYYMMDD_HHMMSS.csv
```

The CSV headers are:

| Email For Account | Calendar Name | Earliest Timestamp | Earliest Name | Latest Timestamp | Latest Name |
|-------------------|--------------|--------------------|---------------|-----------------|-------------|

## How to Use

1. Open the **Script Editor** app on your Mac.
2. Copy and paste the contents of `calendar_event_extremes_to_csv.applescript` into a new document.
3. Press **Run**.
4. When the script completes, a dialog will show you the location of the newly created CSV file.

## Requirements

- macOS with the built-in **Calendar** app.
- Script Editor (included on all Macs).

## Troubleshooting

- If you have a large number of events or calendars, the script may take a few seconds to run.
- The script requires permission to access your Calendars and to write to your Documents folder. Grant access if prompted.

## License

This script is provided as-is, without warranty. You are free to modify or distribute it.
