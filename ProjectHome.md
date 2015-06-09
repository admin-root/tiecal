This utility provides **one way** synchronization between Lotus Notes and Outlook.

From Outlook the calendar can be synchronized to various mobile phones, such as the iPhone or an Android phone. In fact, TieCal was designed to synchronize my Lotus Notes calendar with my iPhone and I use it for this purpose every day.

The source code is open (under the [GNU GPL](http://www.fsf.org/licensing/licenses/gpl.html) license) and there is **no cost** to download or use TieCal.

See ScreenShots for a quick glance at how the tool looks like.

For information how to use it, see UserGuide

## News ##
**Tuesday 2010.03.02:** Released 0.5.1. Changes:
  * Runs on x64 version of Windows (Win7 tested, with Notes 8.5)
  * Fixed rare bug caused by certain items in notes caused tiecal to not sync at all
  * Fixed display of ordinal numbers (1st, 2nd, 3rd) being wrong in merge window

**Monday 2010.01.18:** Released 0.5. Changes:
  * Support for repeating events. Almost all repeating events should now synchronize to Outlook. (The only known exception being "custom" events, which have no repeat pattern)
  * Basic support for synchronizing directly to the iPhone through iTunes.
  * Add a dialog that shows why some entries are skipped during synchronization
  * Numerous bug fixes, including faster startup time

**Sunday 2009.09.27:** Released 0.4. Changes:
  * Entries in outlook (and on iphone) in the "nosync" category will not be touched. This allows you to create entries on the iphone that TieCal won't delete just because they don't exist in notes.
  * Less busy UI: Settings have been moved to the "welcome box" with separate dialog to set DB and reminder settings. Also, it should be more obvious for new users what they need to do before being able to sync
  * Option to skip the "Merge Dialog" that appears before changes are written to outlook.

## DISCLAIMER ##
This tool is provided without an warranty. Backup your Outlook and iPhone calendars before trying this tool.

## Limitations ##
The synchronization only goes one way, from Notes to Outlook. Entries created in Outlook or on your phone will **not** be written to Lotus Notes. I have no plans to add two way sync, since I personally do not need it.

## Planned Improvements ##
See the [Issue Page](http://code.google.com/p/tiecal/issues/list) to view existing feature requests or add your own request.

## Compatibility ##
This tool has been tested with Lotus Notes v7.0.2 and v8.5 and Microsoft Outlook 2007. Other versions may work, but I have no possibility to test this. If you try this, please report success of failure in the [discussion forum](http://groups.google.com/group/tiecal-discussion).

## Requirements ##

### To Compile the Code (Advanced) ###
The tool is written in C# using WPF as GUI toolkit. It uses LINQ to query collections and thus the requirements to compile the sources are:
  * Visual Studio 2008 ([VS2008Express](http://www.microsoft.com/Express/) should be fine)
  * .NET Framework 3.5 (sp1 recommended) ([download](http://www.microsoft.com/downloads/details.aspx?FamilyID=AB99342F-5D1A-413D-8319-81DA479AB0D7&displaylang=en))

### To Run the Program ###
The only unusual runtime dependency is .NET 3.5 runtime (freely available from microsoft, will be distributed when I make a proper setup.exe release). In the meantime, you can download it [from microsoft](http://www.microsoft.com/downloads/details.aspx?familyid=333325FD-AE52-4E35-B531-508D977D32A6&displaylang=en).