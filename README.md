Entourage to iCal export script
===============================

I'm stuck using Microsoft Entourage 2004 at work. I wanted to have access to
my calendar via iCal and on my iPhone (via MobileMe). This script exports the
first Exchange calendar configured in Entourage to `~/Sites/entourage.ics`.
You can then [subscribe to it via iCal][wc], provided you've got Web Sharing
turned on.

If you want access to the calendar via MobileMe, I'm pretty sure that
`entourage.ics` will need to be accessible from the Internet. In that case,
you'll probably want to scp it somewhere, or (maybe?) stick the file in your
iDisk.

INSTALLATION
------------

First, you'll need to [download my fork][fork] of the Python iCalendar module.
Unpack it and then run

    python setup.py install

Next, set the `MY_EMAIL` variable near the top of `export_entourage_cal.py` 
to your email address.

Then, modify `org.bravo5.EntourageExport.plist` and replace the value of
`ProgramArguments` with the path to `export_entourage_cal.py`. Finally, copy
(or symlink) `org.bravo5.EntourageExport.plist` to `~/Library/LaunchAgents/`
and run

    launchctl load ~/Library/LaunchAgents/org.bravo5.EntourageExport.plist

`launchd` will then export your calendar every 15 minutes.

[wc]: webcal://localhost/~USERNAME/entourage.ics
[fork]: https://github.com/blalor/iCalendar/archives/master
