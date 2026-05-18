/*
 * File 1: The JavaScript Engine (commands.js)
 * Purpose: OnMessageSend event handler — checks for external recipients with real attachments.
 *
 * Architecture Notes:
 * - Inline attachments (signature logos, social icons) are filtered out to prevent false positives.
 * - The notificationMessages API is NOT used — we rely exclusively on the Smart Alert dialog
 *   to avoid a UI race condition that can freeze the compose window on some Outlook builds.
 * - All data-gathering API calls (recipients + attachments) run concurrently via Promise.all
 *   to stay well within Microsoft's 5-second execution timeout.
 * - Internal domains are maintained in a single array for easy future expansion.
 */

// ============================================================================
// CONFIGURATION: Add or remove internal domains here as needed.
// ============================================================================
const INTERNAL_DOMAINS = [
    "metrixlab.com",
    "toluna.com"
    // Add future domains below, e.g.:
    // "newentity.com",
];

// ============================================================================
// EVENT HANDLER
// ============================================================================

/**
 * Handles the OnMessageSend event.
 * Runs all API calls concurrently, filters inline attachments,
 * and triggers the Smart Alert pop-up if external recipients have real attachments.
 * @param {Office.AddinCommands.Event} event
 */
function checkExternalAttachments(event) {
    let item = Office.context.mailbox.item;

    // --- Concurrent data gathering (Fix #3: eliminates sequential bottleneck) ---
    let getTo  = new Promise(function(resolve) { item.to.getAsync(function(r)  { resolve(r.value || []); }); });
    let getCc  = new Promise(function(resolve) { item.cc.getAsync(function(r)  { resolve(r.value || []); }); });
    let getBcc = new Promise(function(resolve) { item.bcc.getAsync(function(r) { resolve(r.value || []); }); });
    let getAtt = new Promise(function(resolve) { item.getAttachmentsAsync(function(r) { resolve(r.value || []); }); });

    // Execute ALL four queries simultaneously
    Promise.all([getTo, getCc, getBcc, getAtt]).then(function(results) {
        var toRecipients  = results[0];
        var ccRecipients  = results[1];
        var bccRecipients = results[2];
        var allAttachments = results[3];

        // --- Check for external recipients ---
        var allRecipients = toRecipients.concat(ccRecipients, bccRecipients);
        var externalRecipients = [];

        for (var i = 0; i < allRecipients.length; i++) {
            var emailAddress = allRecipients[i].emailAddress;
            var domain = emailAddress.substring(emailAddress.indexOf("@") + 1).toLowerCase();

            if (INTERNAL_DOMAINS.indexOf(domain) === -1) {
                externalRecipients.push(allRecipients[i].emailAddress);
            }
        }

        var hasExternalRecipients = externalRecipients.length > 0;

        // --- Filter out inline/embedded attachments (Fix #1: eliminates signature false positives) ---
        // Outlook treats embedded signature images (company logos, social media icons)
        // as attachments with isInline = true. We must ignore these.
        var realAttachments = [];
        for (var j = 0; j < allAttachments.length; j++) {
            if (!allAttachments[j].isInline) {
                realAttachments.push(allAttachments[j]);
            }
        }

        var hasRealAttachments = realAttachments.length > 0;

        // --- Decision Logic ---
        if (hasExternalRecipients && hasRealAttachments) {
            // Build a clear, actionable message for the Smart Alert pop-up
            var externalList = externalRecipients.length <= 5
                ? externalRecipients.join(", ")
                : externalRecipients.slice(0, 5).join(", ") + " (+" + (externalRecipients.length - 5) + " more)";

            var alertMessage =
                "⚠ EXTERNAL SEND ALERT\n\n" +
                "You are sending " + realAttachments.length + " attachment(s) to external recipient(s):\n" +
                externalList + "\n\n" +
                "Please verify that you are authorized to share these files outside the organization.";

            // Fix #2: Only use event.completed — no notificationMessages.addAsync call.
            // This avoids the UI race condition that can freeze the compose window.
            event.completed({
                allowEvent: false,
                errorMessage: alertMessage
            });
        } else {
            // No risk detected — allow the email to send normally
            event.completed({ allowEvent: true });
        }

    }).catch(function(error) {
        // Fail-safe: if the API errors out, allow send to avoid blocking business email
        event.completed({ allowEvent: true });
    });
}

// ============================================================================
// FUNCTION ASSOCIATION (Required by Outlook event-based activation)
// ============================================================================
if (Office.actions && Office.actions.associate) {
    Office.actions.associate("checkExternalAttachments", checkExternalAttachments);
}
