# Introduction #
This programs lets you synchronize the calendar in Lotus Notes with Microsoft Outlook. The synchronization goes one way, from lotus to outlook.

Once the calendar is synchronized with Outlook it can be further synchronized with other calendars and mobile phones.

# Configuration #

The configuration area is in the top part of the main window, hidden under an _expander_.

## Notes Section ##
In order to get started synchronizing, you must select which lotus notes database to read calendar entries from. Typically, lotus notes has several different databases for various purposes, but the one you want is probably the same as the one containing your normal emails.

On most systems, the notes database containing the calendar is named: `mail\`_username_`.nsf`

To see a list of databases, press the **Refresh** button. A password prompt will appear where you need to enter your notes email password. From this dialog, you also have the option to save the password for the next time you run `TieCal`.

## Reminders Section ##
The _Reminders_ section controls how reminders are set on calendar entries when they are put in Outlook. The default setting is to not use reminders at all, which means that outlook (or any other device which you later sync with outlook) will **not** warn you before the meeting begins.
You can also tell `TieCal`to use the default setting from Outlook. If you use this option, then you can configure your reminders within Outlook. By default, Outlook will put a reminder 15 minutes before the meeting starts. You can control this from `Tools->Options` in the `Calendar` section in Outlook.
The last reminder option is to set a custom reminder in `TieCal`. If you use this option, all new calendar entries will get the reminder setting you configure here.

# Troubleshooting #
  * Why is `TieCal` removing entries that are created on the phone or in Outlook?
    * This is because `TieCal` only works in one direction: from Notes to Outlook. You can, however (**from `TieCal` version 0.4 and newer**) tell the program to ignore certain entries in Outlook by putting them in a category called "**nosync**"
  * Why can't I click on the Synchronize button? It's dimmed out!
    * Probably you haven't selected which database to read calendar entries from. See the Configuration section above for more details.
  * Why doesn't `TieCal` synchronize all my calendar entries?
    * There could be several reasons. First of all, `TieCal` can only synchronize whatever is in your local database copy - make sure you have replicated from the server before synchronizing. Repeating events were not supported at all in 0.4 and earlier, and there are still some repeated events that aren't supported in the newer versions (typically weird/unusual repeat patterns, like easter which doesn't follow any mathematical formula.)