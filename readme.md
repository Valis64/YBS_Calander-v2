# YBS Print Calander

The YBS Print Calander project provides both a command-line interface (CLI) and a
Tkinter GUI for logging into the YBS portal, confirming the login status, and
retrieving the latest order information. Orders are displayed in a two-column
format showing the order number and the associated company name.

## Features

- **CLI** utility to authenticate and print a formatted order table.
- **Dark blue themed Tkinter GUI** featuring username/password inputs, a
  red/green status light for login feedback, and a live order table.
- Scrapes order numbers and company names from the YBS manage page after a
  successful login.

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

The GUI allows you to enter a username and password, attempt a login, and view
the retrieved orders in a centered two-column table. A yellow indicator shows an
in-progress login, which turns green on success or red on failure.

## Using the CLI

Run the CLI tool to log in and print the order table:

```bash
python -m ybs_print_calander.cli --username YOUR_USERNAME --password YOUR_PASSWORD
```

If you omit either argument, the CLI will prompt interactively. Use
`--no-prompt` to force non-interactive behaviour (in which case both credentials
must be supplied on the command line).

## Notes

- A working internet connection and valid YBS portal credentials are required
  for successful login and order retrieval.
- The project does not store credentials. They are only used for the active
  session with the remote service.
