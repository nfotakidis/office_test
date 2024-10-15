// Ensures the Office.js library is loaded.
Office.onReady();

// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
  // Get the Base64-encoded EML format of a reported message.
  Office.context.mailbox.item.getAsFileAsync(
    { asyncContext: event },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(
          `Error encountered during message processing: ${asyncResult.error.message}`
        );
        return;
      }

      // Get the user's responses to the options and text box in the preprocessing dialog.
      const spamReportingEvent = asyncResult.asyncContext;
      const reportedOptions = spamReportingEvent.options;
      const additionalInfo = spamReportingEvent.freeText;

      // Now, forward the email to a specific recipient
      forwardSpamReport(asyncResult.value, additionalInfo);

      // Signals that the spam-reporting event has completed processing.
      spamReportingEvent.completed({
        onErrorDeleteItem: true,
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
        showPostProcessingDialog: {
          title: "Contoso Spam Reporting",
          description: "Thank you for reporting this message.",
        },
      });
    }
  );
}

// Function to forward the spam report to a specific recipient
function forwardSpamReport(file, additionalInfo) {
  // You can forward the message to a predefined recipient here
  Office.context.mailbox.item.forwardAsync({
    // Define your recipient here
    toRecipients: ["n.fotakidis@kenotom.com"],
    // You can customize the subject of the forwarded message
    subject: "Spam Report: Reported Email",
    // Add any body text you want to include
    body: `A spam report has been submitted.\n\nAdditional Information: ${additionalInfo}`,
  }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to forward spam report: " + asyncResult.error.message);
    } else {
      console.log("Spam report successfully forwarded.");
    }
  });
}

/**
 * IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name
 * specified in the manifest to its JavaScript counterpart.
 */
Office.actions.associate("onSpamReport", onSpamReport);
