var mailboxItem;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

function validateBody(event) {
    mailboxItem.body.getAsync("html",
        { asyncContext: event },
        checkBodyOnlyOnSendCallBack);
}

function validateSubjectAndCC(event) {
    shouldChangeSubjectOnSend(event);
}

function shouldChangeSubjectOnSend(event) {
    mailboxItem.subject.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            addCCOnSend(asyncResult.asyncContext);

            var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)

            subject = '[Checked]: ' + asyncResult.value;

            if (asyncResult.value === null ||
                (/^\s*$/).test(asyncResult.value)) {
                mailboxItem.notificationMessages.addAsync('NoSend',
                    {
                        type: 'errorMessage',
                        message: 'Please enter a subject for this email.'
                    });
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                if (!checkSubject)
                    subjectOnSendChange(subject, asyncResult.asyncContext);
                else
                    asyncResult.asyncContext.completed({ allowEvent: true });
            }
        }
    )
}

function addCCOnSend(event) {
    mailboxItem.cc.setAsync(['roxi@microsoft.com'],
        { asyncContext: event });
}

function subjectOnSendChange(subject, event) {
    mailboxItem.subject.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend',
                    {
                        type: 'errorMessage',
                        message: 'Unable to set the subject.'
                    });
                asyncResult.asyncContext.completed({ allowEvent: false });
            }

            asyncResult.asyncContext.completed({ allowEvent: true });
        });
}

function checkBodyOnlyOnSendCallBack(asyncResult) {
    var listOfBlockedWords = new Array("secret1", "secret2", "secret3");
    var wordExpression = listOfBlockedWords.join('|');
    var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
    var checkBody = regexCheck.test(asyncResult.value);

    if (checkBody) {
        mailboxItem.notificationMessages.addAsync('NoSend',
            {
                type: 'errorMessage',
                message: 'Blocked words have been found in the body of this email. Please remove them.'
            });
        asyncResult.asyncContext.completed({ allowEvent: false });
    }

    asyncResult.asyncContext.completed({ allowEvent: true });
}