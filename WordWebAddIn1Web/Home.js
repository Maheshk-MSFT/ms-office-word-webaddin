
(function () {
    "use strict";
    var messageBanner;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");

                //$('#button-text').text("Display!");
                //$('#button-desc').text("Display the selected text");
                
                $('#highlight-button').click(displaySelectedText);
                return;
            }

            $("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
            //$('#button-text').text("Highlight!");
            //$('#button-desc').text("Highlights the longest word.");

            $('#AddPara1-text').text("Add Para1");
            $('#AddPara1-desc').text("Add Para1- insert new text");
            
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightLongestWord);
            $('#AddPara1button').click(AddPara1);
            $('#GetAtomicNumber').click(GetAtomicNumber);
        });
    };

    function AddPara1() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            //body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "<b>Today was Day 1 </b> of the 2017 Microsoft Ignite conference. There were so many awesome feature announcements made, and TONS of great sessions to attend. Here’s a short summary of what was announced today. I present this as a sort of “mini-Azure Weekly” since there’s so many awesome things to look at today!",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
            .catch(errorHandler);
    }

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

    function hightlightLongestWord() {
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
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
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

    function GetAtomicNumber(elementname) {

        var inputtext = $("#input")[0].value;
        $("#Result2").html("<br> <b> <table><tr></table>");
        if (inputtext == "") {
            app.showNotification("ERROR:", "Input text cannot be null");
        }
        else {
            var codebehindPage = "../temppage.aspx?requesttype=getatomicnumber&input=" + inputtext;
            var xmlhttp;
            try {
                xmlhttp = new XMLHttpRequest();
                xmlhttp.open("GET", codebehindPage, false);
                xmlhttp.send(null);
                if (xmlhttp.status == 200) {
                    var result = xmlhttp.responseText;
                    var myResultArray = result.split("~");
                    var htmlstring = "<table class='hoverTable'>";
                    for (var i = 1; i < myResultArray.length; i++) {
                        var finalValue = myResultArray[i].trim().split(",");
                        htmlstring += "<tr background-color=#ffff99><td>" + finalValue[0] + "</td><td>" + finalValue[1] + "</td></tr>";
                    }
                    htmlstring += "</table><b> </br>";
                    $("#Result2").html(htmlstring);
                    $("#Result2").visible = false;
                }
                if (xmlhttp.status == 405) {
                    app.showNotification("ERROR:", xmlhttp.status);
                }
            }
            catch (ex) {
                app.showNotification("ERROR:", xmlhttp.message);
            }
        }

        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            //body.clear();
            // Queue a command to insert text into the end of the Word document body.

            //var s = "string" + $("#Result2").html;
            
            body.insertText(
                $("#Result2")[0].innerText,
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
            .catch(errorHandler);

    }

})();
