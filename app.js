// Copyright © 2021 Forcepoint LLC. All rights reserved.

let logEnable = true;

function sleep(delay) {
    var start = new Date().getTime();
    while (new Date().getTime() < start + delay);
}

Office.initialize = function () {
}

function printLog(text) {
    if(logEnable){
        Office.context.mailbox.item.notificationMessages.replaceAsync("succeeded", {
            type: "progressIndicator",
            message: text,
        });
      sleep(1500);
    }
}

async function postData(url = '', data = {}, event) {
    printLog("Sending event to classifier");
    const controller = new AbortController()
    const timeout = setTimeout(() => {
        controller.abort();},
        30000)
    fetch(url, {
        signal: controller.signal,
        method: 'POST',
        mode: 'cors',
        cache: 'no-cache',
        credentials: 'same-origin',
        headers: {
            'Content-Type': 'application/json'
        },
        redirect: 'follow',
        referrerPolicy: 'no-referrer',
        body: JSON.stringify(data)
    }).then(response => {
        if (!response.ok) {
            printLog("Engine returned error: "+response.json());
            handleError(response, event);
        }
        return response.json();
    }).then(response => {
        clearTimeout(timeout);
        handleResponse(response, event);
    }).catch(e => {
        printLog("Request crashed");
        handleError(e, event);
    });
}

function handleResponse(data, event) {
    printLog("Handling response from engine");
    message = Office.context.mailbox.item;
    if (data["action"] === 1) {
        message.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked by DLP engine' });
        console.log("DLP block");
        event.completed({ allowEvent: false });
    } else if (data["action"] === 0) {
        printLog("Completing event: DLP Allow");
        console.log("DLP allow");
        event.completed({ allowEvent: true });
    } else {
        printLog("Completing event: Unrecognized JSON response: " + JSON.stringify(data));
        console.log("Unrecognized JSON response: " + JSON.stringify(data));
        event.completed({ allowEvent: true });
    }
}

async function tryPost(event, subject, from, to, cc, bcc, body, attachments) {
        const data = {
            "subject": subject,
            "body": body,
            "from": from,
            "to": to,
            "cc": cc,
            "bcc": bcc,
            "attachments": attachments
        };
        console.log(data);
        postData('https://localhost:55296/OutlookAddin', data, event)

}

async function getAttachment(message, id, result, attachments, post, continuation) {
    printLog("Getting attachment");
    message.getAttachmentContentAsync(id, data => {
        let attachment = {
            "data": data.value.content,
            "content_type": result.contentType,
            "file_name": result.name
        };

        attachments.push(attachment);

        if (post)
            continuation(attachments);
    });
}

async function getAttachmentsList(message, event, subject, from, to, cc, bcc, body) {
    printLog("Getting attachments");
    message.getAttachmentsAsync(result => {
        let attachments = [];
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            if (result.value.length > 0) {
                printLog("Total attachments: "+result.value.length);
                const length = result.value.length;
                for (let i = 0; i < length; i++) {
                    const id = result.value[i].id;
                    const res = result.value[i];
                    const post = i + 1 == length;
                    getAttachment(message, id, res, attachments, post, attachments => {
                        tryPost(event, subject, from, to, cc, bcc, body, attachments);
                    });
                }
            } else {
                printLog("No attachments");
                tryPost(event, subject, from, to, cc, bcc, body, attachments);
            }
        }
    });
}

function getIfVal(result)
{
    if (result.status === Office.AsyncResultStatus.Succeeded) {
        return result.value;
    }
    return "";
}

async function validate(event) {
    message = Office.context.mailbox.item;
    if (message.itemType == "appointment") {
        printLog("Validating appointment");
        message.subject.getAsync(result => {
            let subject = getIfVal(result);
            message.organizer.getAsync(result => {
                let organizer = getIfVal(result);
                message.requiredAttendees.getAsync(result => {
                    let requiredAttendees = getIfVal(result);
                    message.optionalAttendees.getAsync(result => {
                        let optionalAttendees = getIfVal(result);
                        message.body.getAsync("html", {asyncContext: event}, result => {
                            let body = getIfVal(result);
                            getAttachmentsList(message, event, subject, organizer, requiredAttendees, optionalAttendees, [], body);
                        });
                    });
                });
            });
        });
    } else if (message.itemType == "message") {
        printLog("Validating message");
        message.subject.getAsync(result => {
            let subject = getIfVal(result);
            message.from.getAsync(result => {
                let from = getIfVal(result);
                message.to.getAsync(result => {
                    let to = getIfVal(result);
                    message.cc.getAsync(result => {
                        let cc = getIfVal(result);
                        message.bcc.getAsync(result => {
                            let bcc = getIfVal(result);
                            message.body.getAsync("html", {asyncContext: event}, result => {
                                let body = getIfVal(result);
                                getAttachmentsList(message, event, subject, from, to, cc, bcc, body);
                            });
                        });
                    });
                });
            });
        });
    } else {
        console.log("message item type unknown");
        console.log(message.itemType);
        printLog("message item type unknown");
        printLog(message.itemType);
        handleError("Unknown Message Type", event)
    }
}

function handleError(data, event) {
    console.log(data);
    console.log(event);
    printLog("Completing event: "+data);
    event.completed({ allowEvent: true });
    printLog("Event Completed");
}

function operatingSytem() { 
    var contextInfo = Office.context.diagnostics;
    var platform = contextInfo.platform;
    console.log('Office application: ' + contextInfo.host);
    printLog('Office application: ' + contextInfo.host);
    console.log('Platform: ' + platform);
    printLog('Platform: ' + platform);
    console.log('Office version: ' + contextInfo.version);
    printLog('Office version: ' + contextInfo.version);

    //Add code here to write above information to message body as well as console.
    if(platform == 'Mac'){return 'MacOS';}
    if(platform == 'OfficeOnline'){return 'MacOS';}
    if(platform == 'PC'){return 'Other';}
    return 'Other'
} 

function validateBody(event) {
    Office.onReady().then(function() {
        Office.context.mailbox.item.notificationMessages.replaceAsync("succeeded", {
            type: "progressIndicator",
            message: "Microsoft is working on your request.",
        });
        printLog("FP email validation started");
        if(operatingSytem() == "MacOS"){
            printLog("MacOS detected");
            validate(event).catch(data => {handleError(data, event)});
        } else{
            printLog("OS is not MacOS");
            handleError("Not MacOS", event);
        }
    });
}


if (typeof exports !== 'undefined') {
//    printLog("Export defined")
    exports.handleResponse = handleResponse;
    exports.handleError = handleError;
    exports.postData = postData;
    exports.validateBody = validateBody;
    
}
