/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/* eslint-disable office-addins/no-office-initialize */

let mailboxItem: Office.MessageCompose;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
            ? window
            : typeof global !== "undefined"
                ? global
                : undefined;
}

export const g = getGlobal() as any;

// The ui-less functions need to be available in global scope
g.validateEmailAddresses = validateEmailAddresses;

let addressesDialog: Office.Dialog;
function validateEmailAddresses(event) {

    let recipients: Office.EmailAddressDetails[] = [];

    mailboxItem.to.getAsync({ asyncContext: event }, toAddresses => {
        recipients = recipients.concat(toAddresses.value);

        mailboxItem.cc.getAsync({ asyncContext: event }, ccAddresses => {
            recipients = recipients.concat(ccAddresses.value);

            mailboxItem.bcc.getAsync({ asyncContext: event }, bccAddresses => {
                recipients = recipients.concat(bccAddresses.value);
                let event = bccAddresses.asyncContext;

                let anyExternalUser: boolean = recipients.findIndex(r => r.recipientType === "externalUser" || r.recipientType === "other") !== -1;

                if (anyExternalUser) {
                    let url: string = `${window.location.origin}/dialog.html`;

                    Office.context.ui.displayDialogAsync(url, getDialogOptions(), dialogResult => {
                        addressesDialog = dialogResult.value;
                        addressesDialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => receiveMessage(message, event));
                        addressesDialog.addEventHandler(Office.EventType.DialogEventReceived, () => dialogClosed(event));
                    });

                } else {
                    event.completed({ allowEvent: true });
                }

            });
        });
    });
}

function getDialogOptions(): Office.DialogOptions {

    let dialogOptions: Office.DialogOptions;
    // height: height of the dialog as a percentage of the current display. Defaults to 80%. 250px minimum
    // width: width of the dialog as a percentage of the current display. Defaults to 80%. 150px minimum

    if (Office.context.diagnostics.platform === Office.PlatformType.OfficeOnline) { // Browser

        // On Browser the iframe cannot get the dimensions of the parent window since it is on another domain
        dialogOptions = { width: 35, height: 9, displayInIframe: true }; // Browser

    } else { // Desktop
        let fixedWidth = 550;
        let fixedHeight = 160;
        let percentageWidth: number;
        let percentageHeight: number;

        percentageWidth = Math.round(100 * fixedWidth / screen.width);
        percentageHeight = Math.round(100 * fixedHeight / screen.height);

        dialogOptions = { width: percentageWidth, height: percentageHeight, displayInIframe: false }; // Desktop
    }
    return dialogOptions;
}

function receiveMessage(message: any, event: any) {
    addressesDialog.close();
    addressesDialog = null;

    if (message.message === "Send") {
        event.completed({ allowEvent: true });
    } else {
        event.completed({ allowEvent: false });
    }
}

function dialogClosed(event: any) {
    addressesDialog = null;
    event.completed({ allowEvent: false });
}