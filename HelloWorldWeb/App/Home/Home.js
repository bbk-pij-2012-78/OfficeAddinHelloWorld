/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
            $("#writeDataBtn").click(function (event) { writeData(); });
            $("#readDataBtn").click(function (event) { readData(); });
            $("#bindDataBtn").click(function (event) { bindData(); });
            $("#readBoundDataBtn").click(function (event) { readBoundData(); });
            $("#addEventBtn").click(function (event) { addEvent(); });
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    function writeData() {
        Office.context.document.setSelectedDataAsync([["red"], ["green"], ["blue"]], function (asyncResult) {
            if (asyncResult.status === "failed")
                { writeToPage('Error: ' + asyncResult.error.message); }
        });
    }
    function writeToPage(text) {
        document.getElementById('results').innerText = text;
    }

    function readData() {
        Office.context.document.getSelectedDataAsync("matrix", function (asyncResult) {
            if (asyncResult.status === "failed") {
                writeToPage('Error: ' + asyncResult.error.message);
            }
            else {
                writeToPage('Selected data: ' + asyncResult.value);
            }
        });
    }

    function bindData() { 
        Office.context.document.bindings.addFromSelectionAsync("matrix", { 
            id: 'myBinding' }, function (asyncResult) { 
                if (asyncResult.status === "failed") { 
                    writeToPage('Error: ' + asyncResult.error.message); } 
                else { 
                    writeToPage('Added binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id); 
                } 
            }); 
    }

    function readBoundData() { 
        Office.select("bindings#myBinding").getDataAsync({ coercionType: "matrix" }, function (asyncResult) { 
            if (asyncResult.status === "failed") { 
                writeToPage('Error: ' + asyncResult.error.message); 
            } 
            else { 
                writeToPage('Selected data: ' + asyncResult.value); 
            } 
        }); 
    }

    function addEvent() {
        Office.select("bindings#myBinding").addHandlerAsync("bindingDataChanged", myHandler, function (asyncResult) {
            if (asyncResult.status === "failed") {
                writeToPage('Error: ' + asyncResult.error.message);
            } else { writeToPage('Added event handler'); }
        });
    }

    function myHandler(eventArgs) {
        eventArgs.binding.getDataAsync({ coerciontype: "matrix" }, function (asyncResult) {
            if (asyncResult.status === "failed") {
                writeToPage('Error: ' + asyncResult.error.message);
            } else {
                writeToPage('Bound data: ' + asyncResult.value);
            }
        });
    }

})();