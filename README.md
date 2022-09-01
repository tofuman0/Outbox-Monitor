
# Outbox Monitor

An Outlook addin which can monitor the outbox folder for items that have failed to move into the sent items folder after sending.

## So what does this do?

An issue that can present itself when a mailbox is **hosted on 365** is that a sent item may fail to move into the sent items folder after sending.

These items can be confirmed by opening the item in the **"outbox"** folder and you'll find the email has a **"sent on"** date and the **"send" button isn't available**.

A temporary fix can be to recreate the outlook profile but the issue will likely return. **The issue only appears to affect cached mode** so suspect the OST and syncing issues although **no errors** are logged in the "sync issues" folder.

The items can be moved into the "sent items" folder but can become tedious. Especially when working from a unstable internet connection the items "stuck" in the outbox can mount up. As a work around this plugin checks for these items and moves them into the "sent items".

***Items that are in the outbox but haven't been sent aren't processed by this addin. So if you run outlook in offline mode it will ignore any emails that haven't actually sent.***

## Issues

Currently this addin only monitors the default mailbox. So shared mailboxes and additional mailboxes setup in outlook wont be monitored by this addin. Plans are to add support for multiple mailboxes in outlook but the issue only appears to affect the primary mailbox so this currently does provide a solution.

## Configuration

The configuration file is located in **%localappdata%\Outbox Monitor** and is called "Config.json". If the file doesn't exist it will be created.

Example:
```
{
  "BackgroundMonitor": true,
  "BackgroundInterval": 60,
  "LogLevel": 2,
  "LogOnly": false
}
```

### Configuration items:

**BackgroundMonitor:** Set true or false. Enables background thread to check for outbox items.  
**BackgroundInterval:** The **interval in seconds** in which the background monitor checks for outbox items.  
**LogLevel:** The level in which to log to the log files:  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;**0:** Log only errors  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;**1:** Log warnings and errors  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;**2:** Log information, warnings and errors  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;**3:** No logging  
**LogOnly:** Set true or false. Only log the items found in the outbox to the information log file (only one log entry for unique results).  

***Log files are stored in %localappdata%\Outbox Monitor\Logs***