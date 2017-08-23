var mailboxItem;
var mailbox;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
    mailbox = Office.context.mailbox;
    console.log("r: office initialize");
    mailboxItem.loadCustomPropertiesAsync(customPropsCallback);
}

// set customed property
function customPropsCallback(asyncResult) {
    console.log("calling custom property callback");

    var customProps = asyncResult.value;
    var myProp = customProps.get("myProp");

    customProps.set("otherProp", "value");
    customProps.saveAsync(saveCallback);
  }
  
  function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed){
      console.error("r: " + asyncResult.error.message);
    }
    else {
        console.log("r: success");
      // Async call to save custom properties completed.
      // Proceed to do the appropriate for your add-in.
    }
  }

  // Get Changekey from EWS
  function getItemDataRequest(item_id) {
    var request;

    request = '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '    <GetItem' +
        '                xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '      <ItemShape>' +
        '        <t:BaseShape>IdOnly</t:BaseShape>' +
        '      </ItemShape>' +
        '      <ItemIds>' +
        '        <t:ItemId Id="' + item_id + '"/>' +
        '      </ItemIds>' +
        '    </GetItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';

    return request;
}

// Set internet header by calling EWS api, ChangeKey is required
var getRequest = function(id){
    return '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    '  <soap:Header>' +
    '    <t:RequestServerVersion Version="Exchange2013" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +

    '        <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AutoResolve">' +
    '            <m:ItemChanges>' +
    '                <t:ItemChange>' +
    '                    <t:ItemId Id="' + id + '" ChangeKey="CQAAABYAAADEGdy7h2iPRZC3bJfB4OpAAAAAAFKI"/>' +
    '                    <t:Updates>' +
    '                        <t:SetItemField>' +
    '          <t:ExtendedProperty>' +
    '            <t:ExtendedFieldURI DistinguishedPropertySetId="InternetHeaders"' +
    '                                PropertyName="X-OBM"' +
    '                                PropertyType="String" />' +
    '            <t:Value>True</t:Value>' +
    '          </t:ExtendedProperty>' +

    '                        </t:SetItemField>' +
    '                    </t:Updates>' +
    '                </t:ItemChange>' +
    '            </m:ItemChanges>' +
    '        </m:UpdateItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';
}

function sendEWSRequest(){
    
    mailboxItem.saveAsync(function(result){
        itemId = result.value;

        var request = getItemDataRequest(itemId);

        console.log("r: send ews request" + itemId);
        mailbox.makeEwsRequestAsync(request, function(asyncResult){

            console.log("r: "+ asyncResult.value);
            var prop = null;
            try {
                var response = $.parseXML(asyncResult.value);
                var responseDOM = $(response);
    
                console.log(responseDOM);
                if (responseDOM) {
                    prop = responseDOM.filterNode("t:ItemId")[0];
                }
    
            } catch (e) {
                errorMsg = e;
            }
    
            //changeKey = prop.getAttribute("ChangeKey");

            //console.log("r: "+ changeKey);
    
            request = getRequest(itemId);
            mailbox.makeEwsRequestAsync(request, function(asyncResult){
                var response = asyncResult.value;
                var context = asyncResult.context;
                console.log("r: " + response);
            });
            console.log("r: ews request sent");
        });
    });
}

function buttonClick(){
    console.log("r: hello");
    sendEWSRequest();
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