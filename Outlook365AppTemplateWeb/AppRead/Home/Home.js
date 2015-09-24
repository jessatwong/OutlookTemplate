/// <reference path="../App.js" />


var entities = "";

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            entities = Office.context.mailbox.item.getEntities();

            if ($('.ms-FilePicker').length > 0) {
                $('.ms-FilePicker').css({ "position": "absolute !important" });
                $('.ms-Panel').FilePicker();

            } else {
                if ($.fn.Pivot) {
                    $('.ms-Pivot').Pivot();
                }
            }

            injectContent();

            //inject handoff buttons
            injectHandoff();
        });
    };

    function injectContent() {
        $("#messageContainer").text(templateConfig.Branding.Message);
    }

    /*Injects links to webpages or servers with address found in email message body. Can be replaced with other Outlook
    entities that are found and sent to the web server.*/
    function injectHandoff() {
        // Extract entites and format handoff strings
        //var entities = item.getEntities();

        //Format Address and inject handoff forms
        for (var i = 0; i < templateConfig.Handoff.length; i++) {
            var handoffObject = templateConfig.Handoff[i];
            //Checks for URL replacement and replaces it with the address detected
            if (handoffObject.URI.indexOf("{{ADDRESS}}") > -1) {
                handoffObject.URI = handoffObject.URI.replace("{{ADDRESS}}", encodeURI(entities.addresses[0]));
            }
            var handoffLink = "";
            var webLink = "<span class=\"ms-Button-label\">" +
                templateConfig.Branding.ButtonText + " Using " + handoffObject.Platform + " (" + handoffObject.HandoffType + ")</span>";
            //form submit if using POST
            if (handoffObject.HandoffType == "POST") {
                handoffLink = $("<form />").attr({
                    method: "POST",
                    action: handoffObject.URI,
                    target: "_blank"
                });
                var hidden = $("<input />").attr({
                    type: "hidden",
                    name: handoffObject.Address,
                    value: entities.addresses[0]
                });
                handoffLink.append(hidden);
                var btn = $("<button/>").attr({
                    class: "ms-Button",
                    type: "submit",
                });
                btn.append(webLink);
                handoffLink.append(btn);
            } else {
                handoffLink = $("<a />").attr({
                    class: "ms-Button",
                    href: handoffObject.URI,
                    target: "_blank",
                });
                handoffLink.append(webLink + "<br/>");
            }
            //populate the area with event handlers
            if (i < templateConfig.Handoff.length / 2) {
                $("#webLinkListCol1").append(handoffLink);
            } else {
                $("#webLinkListCol2").append(handoffLink);
            }
        }
    }

})();

//TBI Other Entites from https://msdn.microsoft.com/en-us/library/office/fp161071.aspx 

function webLinkToggle(hide) {
    if (hide) {
        $("#overallWebLink").hide();
    } else {
        $("#overallWebLink").show();
    }
    $("#displayEntities").empty();
}

/*Displays all the information about the contacts mentioned in the email message body if found. Contacts must be referenced
eiither as a signature or the name must be mentioned within the vicinity of phone numbers, businesses, email address, or URLs.
Otherwise, users are told that no contacts were found. */
function getContacts() {
    webLinkToggle(true);
    //all contacts found in message
    var contactsArray = entities.contacts;
    if (contactsArray != null && contactsArray.length > 0) {
        var contactsHTML = "";
        for (var i = 0; i < contactsArray.length; i++) {
            //Name of contact
            var name = "<b>" + contactsArray[i].personName + " </b><br/>";

            //Name of business associated with contact
            var business = "<b>Business: </b>" + contactsArray[i].businessName + " <br/>";

            //List of phone numbers associated with contact if found
            var phoneNumberArray = contactsArray[i].phoneNumbers;
            var phoneList = "";
            if (phoneNumberArray != null) {
                for (var j = 0; j < phoneNumberArray.length; j++) {
                    phoneList += "<b>Phone: </b>" + phoneNumberArray[j].phoneString + "<br/>";
                    phoneList += "<b>Original Phone: </b>" + phoneNumberArray[j].originalPhoneString + "<br/>";
                }
            }

            //List of URLs associated with contact if found
            var urlArray = contactsArray[i].urls;
            var urlList = "";
            if (urlArray != null) {
                for (var j = 0; j < urlArray.length; j++) {
                    urlList += "<b> URL: </b>" + urlArray[j] + "<br/>";
                }
            }

            //List of Emails associated with contact if found
            var emailArray = contactsArray.emailAddresses;
            var emailList = "";
            if (emailArray != null) {
                for (var j = 0; j < emailArray.length; j++) {
                    emailList += "<b> Email: </b>" + emailArray[j] + "<br/>";
                }
            }
            contactsHTML += name + business + phoneList + urlList + emailList;
        }
        $("#displayEntities").append(contactsHTML);
    } else {
        $("#displayEntities").append("No contacts found in message. Please send contacts as: " +
            "<br/> John Smith <br/> 1075 La Avenida St <br/> jsmith@microsoft.com <br/> 800-123-4567");
    }
}

/*Display all the information about the meeting suggestions mentioned in the message. Meeting suggestions depend on recognizing
events or meetings in the email message. An example of a phrase it will recognize is, "Let's meet [time] for [event or meeting]"*/
function getMeetings() {
    webLinkToggle(true);
    var meetingArray = entities.meetingSuggestions;
    if (meetingArray != null && meetingArray.length > 0) {
        var meetingHTML = "";
        for (var i = 0; i < meetingArray.length; i++) {
            var meeting = meetingArray[i];
            var meetingName = "<b>Meeting String: </b>" + meeting.meetingString + "<br/>";
            var meetingAttendeesArray = meeting.attendees;
            var meetingAttendeesList = "<b>Attendees: </b><br/>";
            if (meetingAttendeesArray != null) {
                for (var j = 0; j < meetingAttendeesArray.length; j++) {
                    meetingAttendeesList += "Display Name: " + meetingAttendeesArray[j].displayName + "<br/>";
                    meetingAttendeesList += "Email Address: " + meetingAttendeesArray[j].emailAddress + "<br/><br/>"
                }
            }
            var meetingLocation = "<b>Location: </b>" + meeting.location + "<br/>";
            var meetingSubject = "<b>Subject: </b>" + meeting.subject + "<br/>";
            var meetingStart = "<b>Start time: </b>" + meeting.start + "<br/>";
            var meetingEnd = "<b>End time: </b>" + meeting.end + "<br/>";

            meetingHTML += meetingName + meetingAttendeesList + meetingLocation + meetingSubject + meetingStart + meetingEnd;
        }
        $("#displayEntities").append(meetingHTML);
    } else {
        $("#displayEntities").append("No meeting suggestions found. Please phrase meeting suggestion as: " +
            "<br/><i>Let's meet [time] for [event or meeting]</i>");
    }
}

/*Display task suggestions mentioned in the message. Task suggestions are recognized as actionable sentences,
i.e "Please update the spreadsheet." Otherwise, users are told that no task suggestions were found. */
function getTasks() {
    webLinkToggle(true);
    var taskArray = entities.taskSuggestions;
    if (taskArray != null && taskArray.length > 0) {
        var taskHTML = "";
        for (var i = 0; i < taskArray.length; i++) {
            var task = taskArray[i];
            var taskName = "<b>Task String: </b>" + task.taskString + "<br/>";

            var taskAssignArray = task.assignees;
            var taskAssignList = "<b>Assignees: </b>";
            if (taskAssignArray != null) {
                for (var j = 0; j < taskAssignArray.length; j++) {
                    taskAssignList += "Display Name: " + taskAssignArray[j].displayName + "<br/>";
                    taskAssignList += "Email Address: " + taskAssignArray[j].emailAddress + "<br/>";
                }
            }
            taskHTML += taskName + taskAssignList;
        }
        $("#displayEntities").append(taskHTML);
    } else {
        $("#displayEntities").append("No task suggestions found. Please phrase task as an actionable sentence");
    }
}


//TBI Custom Entities         

