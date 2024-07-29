function approveLeave(requestId) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests");
    if (!sheet) {
        Logger.log("Sheet 'LeaveRequests' not found.");
        return;
    }

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] == requestId) { // Assuming requestId is in the first column
            sheet.getRange(i + 1, 7).setValue("Approved"); // Assuming Status is in the 7th column
            sendNotification(data[i][1], "Your leave request has been approved.");
            Logger.log("Leave request " + requestId + " approved.");
            break;
        }
    }
}

function rejectLeave(requestId) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests");
    if (!sheet) {
        Logger.log("Sheet 'LeaveRequests' not found.");
        return;
    }

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        if (data[i][0] == requestId) { // Assuming requestId is in the first column
            sheet.getRange(i + 1, 7).setValue("Rejected"); // Assuming Status is in the 7th column
            sendNotification(data[i][1], "Your leave request has been rejected.");
            Logger.log("Leave request " + requestId + " rejected.");
            break;
        }
    }
}

function sendNotification(email, message) {
    if (!email || !message) {
        Logger.log("Invalid email or message.");
        return;
    }
    try {
        MailApp.sendEmail(email, "Leave Request Notification", message);
        Logger.log("Notification sent to: " + email);
    } catch (e) {
        Logger.log("Failed to send email to: " + email + " Error: " + e.message);
    }
}

function sendUpcomingLeaveReminders() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LeaveRequests");
    if (!sheet) {
        Logger.log("Sheet 'LeaveRequests' not found.");
        return;
    }

    var data = sheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0); // Ensure we are comparing dates only
    for (var i = 1; i < data.length; i++) {
        var startDate = new Date(data[i][4]); // Assuming Start Date is in the 5th column
        startDate.setHours(0, 0, 0, 0); // Ensure we are comparing dates only
        var diffDays = Math.ceil((startDate - today) / (1000 * 60 * 60 * 24));
        if (diffDays <= 2 && data[i][6] == "Approved") { // 2 days before leave and status is Approved
            sendNotification(data[i][1], "Reminder: Your leave starts in " + diffDays + " days.");
            Logger.log("Reminder sent to: " + data[i][1] + " for leave starting in " + diffDays + " days.");
        }
    }
}
