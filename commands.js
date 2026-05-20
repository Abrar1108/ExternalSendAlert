/*
 * External Send Alert — commands.js  (v6.0)
 *
 * Modeled exactly after Microsoft's official Smart Alerts sample code.
 * Uses pure callback pattern (no Promises) for maximum compatibility.
 */

// ============================================================================
// CONFIGURATION
// ============================================================================
var INTERNAL_DOMAINS = [
    "metrixlab.com",
    "toluna.com"
];

// ============================================================================
// HELPER: Check if domain is internal (supports subdomains)
// ============================================================================
function isInternalDomain(domain) {
    for (var i = 0; i < INTERNAL_DOMAINS.length; i++) {
        var d = INTERNAL_DOMAINS[i];
        if (domain === d || (domain.length > d.length && domain.indexOf("." + d) === domain.length - d.length - 1)) {
            return true;
        }
    }
    return false;
}

// ============================================================================
// HELPER: Extract external emails from a recipient array
// ============================================================================
function getExternalsFromList(recipients) {
    var externals = [];
    for (var i = 0; i < recipients.length; i++) {
        var email = recipients[i].emailAddress || "";
        if (!email || email.indexOf("@") === -1) continue;
        var domain = email.substring(email.indexOf("@") + 1).toLowerCase();
        if (!isInternalDomain(domain)) {
            externals.push(email);
        }
    }
    return externals;
}

// ============================================================================
// HELPER: Format list for display
// ============================================================================
function formatList(arr) {
    if (arr.length <= 3) return arr.join(", ");
    return arr.slice(0, 3).join(", ") + " (+" + (arr.length - 3) + " more)";
}

// ============================================================================
// OnMessageSend handler — uses asyncContext callback chain (Microsoft pattern)
// ============================================================================
function onMessageSendHandler(event) {
    var item = Office.context.mailbox.item;

    // Step 1: Get To recipients
    item.to.getAsync({ asyncContext: { event: event, all: [] } }, function (toResult) {
        var ctx = toResult.asyncContext;
        var toList = toResult.value || [];
        ctx.all = ctx.all.concat(toList);

        // Step 2: Get Cc recipients
        item.cc.getAsync({ asyncContext: ctx }, function (ccResult) {
            var ctx2 = ccResult.asyncContext;
            var ccList = ccResult.value || [];
            ctx2.all = ctx2.all.concat(ccList);

            // Step 3: Get Bcc recipients
            item.bcc.getAsync({ asyncContext: ctx2 }, function (bccResult) {
                var ctx3 = bccResult.asyncContext;
                var bccList = bccResult.value || [];
                ctx3.all = ctx3.all.concat(bccList);

                var externals = getExternalsFromList(ctx3.all);

                // No external recipients — allow send immediately
                if (externals.length === 0) {
                    ctx3.event.completed({ allowEvent: true });
                    return;
                }

                // External recipients found — check attachments
                if (typeof item.getAttachmentsAsync === "function") {
                    item.getAttachmentsAsync({ asyncContext: { event: ctx3.event, externals: externals } }, function (attResult) {
                        var ctx4 = attResult.asyncContext;
                        var attachments = attResult.value || [];
                        var realCount = 0;
                        for (var j = 0; j < attachments.length; j++) {
                            if (!attachments[j].isInline) realCount++;
                        }
                        blockWithMessage(ctx4.event, ctx4.externals, realCount);
                    });
                } else {
                    blockWithMessage(ctx3.event, externals, 0);
                }
            });
        });
    });
}

// ============================================================================
// Build alert message and block send
// ============================================================================
function blockWithMessage(event, externals, attachmentCount) {
    var msg = "EXTERNAL SEND ALERT\n\nExternal recipient(s): " + formatList(externals) + "\n\n";
    if (attachmentCount > 0) {
        msg += "Attachments: " + attachmentCount + " file(s) included.\n\n";
    }
    msg += "Verify you are authorized to send outside the organization.";

    if (msg.length > 240) {
        msg = msg.substring(0, 237) + "...";
    }

    event.completed({
        allowEvent: false,
        errorMessage: msg
    });
}

// ============================================================================
// Ribbon button handler
// ============================================================================
function onButtonClickHandler(event) {
    var item = Office.context.mailbox.item;

    item.to.getAsync(function (toResult) {
        var toList = toResult.value || [];
        item.cc.getAsync(function (ccResult) {
            var ccList = ccResult.value || [];
            item.bcc.getAsync(function (bccResult) {
                var bccList = bccResult.value || [];
                var all = toList.concat(ccList, bccList);
                var externals = getExternalsFromList(all);

                if (externals.length > 0) {
                    item.notificationMessages.replaceAsync("externalWarning", {
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: "External recipients: " + formatList(externals),
                        icon: "Icon.16x16",
                        persistent: false
                    });
                } else {
                    item.notificationMessages.replaceAsync("externalClear", {
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: "All recipients are internal.",
                        icon: "Icon.16x16",
                        persistent: false
                    });
                }
                event.completed();
            });
        });
    });
}

// ============================================================================
// IMPORTANT: Map event handler names from manifest to JS functions.
// Must be at top level, NOT inside Office.onReady().
// ============================================================================
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onButtonClickHandler", onButtonClickHandler);