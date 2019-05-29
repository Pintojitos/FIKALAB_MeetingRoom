import win32
import win32com.client
import datetime
import dateutil.parser
import time

GLOBAL_SNOOZE_TIME = 5

import codecs

def sendwarningmail(eventSubject, eventOwner, eventID):
    print("Sending warning email!")

    outlook = win32com.client.Dispatch("outlook.application")

    mail = outlook.CreateItem(0)
    mail.To = eventOwner
    mail.Subject = 'Room reservation for event ' + eventSubject
    mail.Body = "Please make sure you still need your room reservation for "+ eventSubject + ".\n\nIf you no longer require the room, delete or cancel the appointment.\n\nIn case no consistent movement is detected for the first 15 minutes of the room's reservation, the room will become vacant for other users.\n\nAny further questions report to: ..."
    mail.SentOnBehalfOfName = "pdpinto@criticalsoftware.com" # Use then the email of the room

    #######################################
    ##  TO DO
    #######################################

    # - Introduce a hyperlink for the event
    # - Introduce an automatic response with mailto in order to send an email to all participants
    # - More attractive text

    #######################################
    ##  TO DO - Difficult
    #######################################
    # - Snooze button
    # - Instant email reply to cancel the meeting (it has to be a readable email template

    #f = codecs.open("buttt.html", 'r', 'utf-8')
    #mail.HTMLBody = f.read() # this field is optional

    # To attach a file to the email (optional):
    # attachment = "Path to the attachment"
    # mail.Attachments.Add(attachment)

    mail.Send()

def sendOrganizerCancelMail(eventSubject, eventOrganizer):
    outlook = win32com.client.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    print("Event Organizer: "+ str(eventOrganizer))
    mail.To = 'pedrodamaspinto@gmail.com' # Set this afterwards eventOrganizer
    mail.Subject = 'Room reservation for event ' + eventSubject + ' is now cancelled.'
    mail.Body = "Your room reservation for " + eventSubject + " as no movement was detected inside the room and it has passed 15 minutes of the start meeting time.\n\nAny further questions report to: ..."
    mail.SentOnBehalfOfName = "pdpinto@criticalsoftware.com"  # Use then the email of the room

    #######################################
    ##  TO DO
    #######################################

    # - Introduce a hyperlink for the event
    # - Introduce an automatic response with mailto in order to send an email to all participants
    # - More attractive text

    #######################################
    ##  TO DO - Difficult
    #######################################
    # - Snooze button
    # - Instant email reply to cancel the meeting (it has to be a readable email template

    # f = codecs.open("buttt.html", 'r', 'utf-8')
    # mail.HTMLBody = f.read() # this field is optional

    # To attach a file to the email (optional):
    # attachment = "Path to the attachment"
    # mail.Attachments.Add(attachment)

    mail.Send()

#Develop this method further
def sendParticipantsCancelMail(eventSubject, eventParticipant):
    print(eventSubject, eventParticipant)

#Not necessary
def addevent(start, subject):
    oOutlook = win32com.client.Dispatch("Outlook.Application")
    appointment = oOutlook.CreateItem(1)  # 1=outlook appointment item
    appointment.Start = start
    appointment.Subject = subject
    appointment.Duration = 20
    appointment.Location = 'Sprortground'
    appointment.ReminderSet = True
    appointment.ReminderMinutesBeforeStart = 1
    appointment.MeetingStatus = 1
    appointment.Recipients.Add("pdpinto@criticalsoftware.com")
    appointment.Send()
    appointment.Save()
    return

#Add a if statement to get only valid appointments (check appointment status)
def getCalendarEntries(appointment):
    # Get all the entries for the day and returns the list of events and the appointment object
    appointment.Sort("[Start]")
    appointment.IncludeRecurrences = "True"
    today = datetime.datetime.today()
    begin = today.date().strftime("%m/%d/%Y")
    tomorrow = datetime.timedelta(days=1) + today
    end = tomorrow.date().strftime("%m/%d/%Y")

    appointment_obj = appointment.Restrict("[Start] >= '" + begin + "' AND [END] <= '" + end + "'")
    events = {'Start': [], 'Subject': [], 'Duration': [], 'Organizer':[], 'Global Appointment ID':[], 'Required Attendees':[]}
    #print(events)
    i = 0
    for a in appointment_obj:
        i = i+1
        # https://docs.microsoft.com/pt-pt/office/vba/api/outlook.appointmentitem.forceupdatetoallattendees
        # For more properties in appointment_obj
        adate = dateutil.parser.parse(str(a.Start))
        print ("Event "+ str(i) +": \n"+ "   Start: \t\t"+str(a.Start) + "\n   Subject: \t" + str(a.Subject) + "\n   Duration: \t"+ str(a.Duration) + "\n   Organizer:\t"+ str(a.Organizer) + "\n      Attendees:\t" + str(a.RequiredAttendees))
        events['Start'].append(str(adate))
        events['Subject'].append(a.Subject)
        events['Duration'].append(a.Duration)
        events['Organizer'].append(a.Organizer)
        events['Global Appointment ID'].append(a.GlobalAppointmentID)
        events['Required Attendees'].append(a.RequiredAttendees)
    return events, appointment_obj



def checkIfStarted(event,app):

    today_now = datetime.datetime.today().timestamp()

    for x in app:
        today_event = x.Start.timestamp()-3600 #3600 to correct the 1 hour add from solar time UTC

        if today_event + x.Duration*60 <= today_now:
            # Passed events
            # Somehow delete the previous events --- or not
            pass

        if today_event <= today_now <= today_event + x.Duration*60:
            #Check for movement and then send email
            noMovement = 1 # if no movement detected -> noMovement = 1
            print("Check for movement.\nIf no movement is detected, send an email")
            if noMovement:
                if today_event + GLOBAL_SNOOZE_TIME * 60 <= today_now: #More than 15 minutes have passed
                    # Send email cancelling the meeting
                    print("************************************************")
                    print("***** Cancelling Event: " + x.Subject + "*******")
                    print("************************************************")
                    print(x.MeetingStatus)
                    sendOrganizerCancelMail(x.Subject, x.Organizer)
                    ###############################################
                    ## Work In Progress - Change the event duration
                    # https://docs.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.propertychange
                    # https://docs.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.meetingstatus
                    # https: // github.com / asmaier / roomfinder / blob / master / README.md


                    x.Duration = 2 #Like this it changes only the meeting duration for this account
                    #x.PropertyChange(x.Duration)
                    x.MeetingStatus=5 # https://docs.microsoft.com/pt-pt/office/vba/api/outlook.olmeetingstatus
                    reqAttendees = x.RequiredAttendees.split('; ')
                    for participants in range(0,len(x.RequiredAttendees.split('; '))):
                        sendParticipantsCancelMail(x.Subject, reqAttendees[participants])

                    x.Send()
                    x.Save()


                    ##############################################
                else:                                                  # Less than 15 minutes have passed
                    print("Sending out email")
                    sendwarningmail(x.Subject, x.SendUsingAccount, x.GlobalAppointmentID)

            print("End of if in CheckIfStarted")
            print(x.Subject)








if __name__=="__main__":

    table = {"4-27": 123456}

    for item in table.keys():
        start = '2019-' + item + ' 18:35'
        subject = 'Experiment. To do:' + str(table[item])
        #addevent(start, subject)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9) # https://docs.microsoft.com/pt-pt/office/vba/api/outlook.oldefaultfolders
    appointments = calendar.Items


    ############################################################
    ### Get all Calendar entries for the day        ############
    ############################################################
    events, appointmentObject = getCalendarEntries(appointments)
    #print("Events:\nStart: " + str(events['Start']) + "\nSubject: " + str(events['Subject']) + "\nDuration: " + str(events['Duration']) + "\n")

    signalNoPeople = 1


    ###############################################################
    ##  TO DO   - please define a better and more readable workflow
    ###############################################################
    # - Set a different flow of the methods
    #      - Get entries for the day
    #      - Check if it is required to check for movement
    #      - Check if it is occupied
    #      - Send email if it is not occupied

    checkIfStarted(events, appointmentObject)



