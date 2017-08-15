var config;

Office.initialize = function (reason) {
    config = getConfig();
};

function showError(error) {
    Office.context.mailbox.item.notificationMessages.replaceAsync('github-error', {
        type: 'errorMessage',
        message: error
    },
        function (result) {
        });
}

var settingsDialog;

function insertDefaultGist(event) {
    if (config && config.defaultGistId) {
        try {
            getGist(config.defaultGistId, function (gist, error) {
                if (gist) {
                    buildBodyContent(gist, function (content, error) {
                        if (content) {
                            Office.context.mailbox.item.body.setSelectedDataAsync(content,
                                { coercionType: Office.CoercionType.Html }, function (result) {
                                    event.completed();
                                });
                        } else {
                            showError(error);
                            event.completed();
                        }
                    });
                } else {
                    showError(error);
                    event.completed();
                }
            });
        } catch (err) {
            showError(err);
            event.completed();
        }
    }
    else {
        //var url = new URI('https://localhost:44364/settings/dialog.html?warn=1').absoluteTo(window.location).toString();
        var url = 'https://localhost:44364/settings/dialog.html?warn=1';
        var dialogOptions = { width: 20, height: 40 };

        Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
            settingsDialog = result.value;
            settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
            settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
            event.completed();
        });
    }
}

function receiveMessage(message) {
    config = JSON.parse(message.message);

    setConfig(config, function (result) {
        settingsDialog.close();
        settingsDialog = null;
    });
}

function dialogClosed(message) {
    settingsDialog = null;
}