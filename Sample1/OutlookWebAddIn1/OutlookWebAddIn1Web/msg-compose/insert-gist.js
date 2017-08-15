(function () {
    'use strict';

    var config;
    var settingsDialog;

    Office.initialize = function (reason) {
        config = getConfig();

        jQuery(document).ready(function () {
            if (config && config.gitHubUserName) {
                loadGists(config.gitHubUserName);
            } else {
                $('#not-configured').show();
            }

            $('#insert-button').on('click', function () {
                var gistId = $('.ms-ListItem.is-selected').children('.gist-id').val();

                getGist(gistId, function (gist, error) {
                    if (gist) {
                        buildBodyContent(gist, function (content, error) {
                            if (content) {
                                Office.context.mailbox.item.body.setSelectedDataAsync(content,
                                    { coercionType: Office.CoercionType.Html },
                                    function (result) {
                                        if (result.status == 'failed') {
                                            showError('Could not insert Gist: ' + result.error.message);
                                        }
                                    });
                            }
                            else {
                                showError('Could not create insertable content: ' + error);
                            }
                        });
                    }
                    else {
                        showError('Could not retreive Gist: ' + error);
                    }
                });
            });

            $('#settings-icon').on('click', function () {
                // Display settings dialog
                //var url = new URI('../settings/dialog.html').absoluteTo(window.location).toString();
                var url = 'https://localhost:44364/settings/dialog.html';
                if (config) {
                    url = url + '?gitHubUserName=' + config.gitHubUserName + '&defaultGistId=' + config.defaultGistId;
                }

                var dialogOptions = { width: 20, height: 40 };

                Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
                    settingsDialog = result.value;
                    settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
                    settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
                    event.completed();
                });
            })
        });
    };

    function loadGists(user) {
        $('#error-display').hide();
        $('#not-configured').hide();
        $('#gist-list-container').show();

        getUserGists(user, function (gists, error) {
            if (error) {

            } else {
                buildGistList($('#gist-list'), gists, onGistSelected);
            }
        });
    }

    function onGistSelected() {
        $('.ms-ListItem').removeClass('is-selected');
        $(this).addClass('is-selected');
        $('#insert-button').removeAttr('disabled');
    }

    function showError(error) {
        $('#not-configured').hide();
        $('#gist-list-container').hide();
        $('#error-display').text(error);
        $('#error-display').show();
    }

    function receiveMessage(message) {
        config = JSON.parse(message.message);
        setConfig(config, function (result) {
            settingsDialog.close();
            settingsDialog = null;
            loadGists(config.gitHubUserName);
        });
    }

    function dialogClosed(message) {
        settingsDialog = null;
    }
})();