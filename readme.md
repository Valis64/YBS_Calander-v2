# YBS Print Calander

The YBS Print Calander project provides both a command-line interface (CLI) and a
Tkinter GUI for logging into the YBS portal, confirming the login status, and
retrieving the latest order information. Orders are displayed in a two-column
format showing the order number and the associated company name.

## Features

- **Command-line interface (CLI)** for authenticating against the YBS portal and
  exporting the order list as a formatted table, CSV, or JSON file directly from
  the terminal.
- **Menu-driven Tkinter GUI** with File/Edit/Settings/Help menus. Use File ▸
  Exit to close the window, the Edit menu (or standard Ctrl/Cmd+Z, Ctrl/Cmd+
  Shift+Z, and Ctrl/Cmd+Y shortcuts) for Undo/Redo, the Settings menu to jump
  between the Orders & Calendar and Settings tabs, and Help ▸ About to see the
  current app version.
- **Settings tab for authentication** that centralizes username and password
  entry, Login and Refresh controls, a color-coded status light (yellow while a
  request is in progress, green on success, red on failure), and a status
  message with the latest refresh timestamp.
- **Orders & Calendar workspace** featuring a filterable orders table,
  drag-and-drop scheduling into a monthly calendar, persistent per-day notes,
  Delete-driven removal of the active day’s assignments, a double-click day
  details pop-up for reviewing and clearing scheduled orders, and Undo/Redo
  support for schedule and note edits.

## Requirements

Install the dependencies before running the application:

```bash
pip install -r requirements.txt
```

## Running the GUI

Launch the GUI directly with:

```bash
python -m ybs_print_calander
```

After launching the window you will land on the **Orders & Calendar** tab. Open
the **Settings** tab (or choose **Settings ▸ Show Settings**) to authenticate:

1. Enter your YBS credentials, then press **Login**. The circular status light
   turns yellow while the request is running, switches to green on success, and
   red on failure. The adjacent status text explains the outcome and the “Last
   updated” timestamp records when orders were most recently downloaded. The
   **Refresh** button remains disabled until a login succeeds, then lets you
   re-fetch orders later without re-entering credentials.
2. Return to **Orders & Calendar** (via the tab header or **Settings ▸ Show
   Orders & Calendar**) to work with the data. The upper filter box narrows the
   orders table in real time. Drag one or more highlighted orders onto a day to
   schedule them, type notes directly into a day’s text area, and press
   **Previous/Today/Next** to navigate between months. Double-click a day cell,
   its notes field, or a scheduled order to open the day details pop-up where
   you can remove selected orders, clear the entire day, or close the dialog.
   Use the **Delete** key to clear the active day (when its header is focused)
   or to remove selected assignments from a day’s order list. Undo and Redo are
   always available through the Edit menu or the standard keyboard shortcuts.
3. The window title displays the semantic version (for example,
   `YBS Print Calander v0.1.0`). Selecting **Help ▸ About** shows the same
   version in an information dialog and echoes it in the Settings tab status
   message so you can confirm which build is running. Use **File ▸ Exit** to
   quit when you are done.

## Using the CLI

Run the CLI tool to log in and print the order table:

```bash
python -m ybs_print_calander.cli --username YOUR_USERNAME --password YOUR_PASSWORD
```

If you omit either argument, the CLI will prompt interactively. Use
`--no-prompt` to force non-interactive behaviour (in which case both credentials
must be supplied on the command line).

The CLI can emit the orders in different formats using `--format`, choosing
between the default `table`, `csv`, or `json`. Combine this with `--output` to
write the formatted results to disk instead of standard output, for example:

```bash
python -m ybs_print_calander.cli --format csv --output orders.csv
```

## Notes

- A working internet connection and valid YBS portal credentials are required
  for successful login and order retrieval.
- The project does not store credentials. They are only used for the active
  session with the remote service.
