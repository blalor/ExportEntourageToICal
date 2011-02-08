#!/usr/bin/python

import icalendar
assert "2.2bjl" in icalendar.__file__, "wrong icalendar; use mine: https://github.com/blalor/iCalendar"

import Cocoa
import Foundation
import PyObjCTools.AppHelper
import ScriptingBridge

import os

## change this to the email address for your Entourage account
MY_EMAIL = "user@example.com"

# enum EntourageEFBt {
#         EntourageEFBtBusy = 'eSBu' /* busy */,
#         EntourageEFBtFree = 'eSFr' /* free */,
#         EntourageEFBtTentative = 'eSTe' /* tentative */,
#         EntourageEFBtOutOfOffice = 'eSOO' /* out of office */
# };

def enum_to_int(enum):
    # http://stackoverflow.com/questions/548289/what-is-the-type-of-an-enum-whose-values-appear-to-be-strings/548305#548305
    return ord(enum[0]) << 24 | ord(enum[1]) << 16 | ord(enum[2]) << 8 | ord(enum[3])


FREE_BUSY_STATUS = {
    enum_to_int('eSBu') : "busy",
    enum_to_int('eSFr') : "free",
    enum_to_int('eSTe') : "tentative",
    enum_to_int('eSOO') : "out of office",
}


app = ScriptingBridge.SBApplication.applicationWithURL_(Foundation.NSURL.fileURLWithPath_("/Applications/Microsoft Office 2004/Microsoft Entourage"))
exchangeAccount = app.ExchangeAccounts()[0]

ical = None

for cal in exchangeAccount.calendars():
    for evt in cal.events():
        # print evt.subject(), evt.startTime(), FREE_BUSY_STATUS[evt.freeBusyStatus()]
        
        ## correct illegal (according to iCal) property name
        ## iCal[63106] <Warning>: Invalid char _ for PropertyName in line 1518
        _ical_data = evt.iCalData()
        _ical_data = _ical_data.replace("X-ENTOURAGE_UUID", "X-ENTOURAGE-UUID")
        
        c = icalendar.Calendar.from_string(_ical_data.encode('utf-8'))
        
        if ical == None:
            ical = c
        
        for component in c.subcomponents:
            if component not in ical.subcomponents:
                if type(component) == icalendar.Event:
                    
                    ## find my attendee entry and set my status to ACCEPTED
                    if 'attendee' in component:
                        attendees = component['attendee']
                        if type(attendees) != type([]):
                            attendees = [attendees]
                            
                        for attendee in attendees:
                            if MY_EMAIL in attendee.lower():
                                if FREE_BUSY_STATUS[evt.freeBusyStatus()] in ("free", "tentative"):
                                    attendee.params['PARTSTAT'] = "TENTATIVE"
                                else:
                                    attendee.params['PARTSTAT'] = "ACCEPTED"
                
                ical.add_component(component)
            
        
    
ofp = open(os.path.expanduser("~/Sites/entourage.ics"), "wb")
ofp.write(ical.as_string())
ofp.close()
