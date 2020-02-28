
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Run now!");
                
                $('#highlight-button').click(displaySelectedText);
                return;
            }

            $("#template-description").text("Select your code and programming language then click 'Run now!'");
            $('#button-text').text("Run now!");
            
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(runCode);
        });
    };

    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
        .catch(errorHandler);
    }

    function runCode() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();
            
            // This variable will keep the search results for the longest word.
            var searchResults;
            
            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    return new Promise(function (resolve, reject) {
                        var code = range.text;
                        let language = document.getElementById("select-languages").value.toString();
                        let stdIn = document.getElementById("input-input").value;
                        fetch("https://codeathon.itsslab.xyz/compile", {
                            method: 'POST', // *GET, POST, PUT, DELETE, etc.
                            mode: 'cors', // no-cors, *cors, same-origin
                            cache: 'no-cache', // *default, no-cache, reload, force-cache, only-if-cached
                            credentials: 'same-origin', // include, *same-origin, omit
                            headers: {
                                'Content-Type': 'application/json'
                                // 'Content-Type': 'application/x-www-form-urlencoded',
                            },
                            redirect: 'follow', // manual, *follow, error
                            referrerPolicy: 'no-referrer', // no-referrer, *client
                            body: JSON.stringify({
                                language: language,
                                code: code,
                                stdin: stdIn
                            }) // body data type must match "Content-Type" header
                        }).then(function (response) {
                            return response.json();
                        }).then(function (json) {
                            var body = context.document.body;
                            if (json.errors) {
                                showNotification("Error", json.errors);
                                body.insertText("\n----------", Word.InsertLocation.end);
                                body.insertText("Errors: " + json.errors, Word.InsertLocation.end);
                                body.insertText("----------\n", Word.InsertLocation.end);
                            } else {
                                showNotification("Run result after " + json.time + "s", json.output);
                                body.insertText("\n----------\n", Word.InsertLocation.end);
                                body.insertText("Runtime: " + json.time + " Result:\n", Word.InsertLocation.end);
                                body.insertText(json.output, Word.InsertLocation.end);
                                body.insertText("\n----------\n", Word.InsertLocation.end);
                            }
                            resolve(json);
                        });
                    }); 
                }).then(context.sync);
        })
        .catch(errorHandler);
    } 


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
