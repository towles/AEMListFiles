(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            //messageBanner.hideBanner();
            
            
            $('#get-data-from-selection').click(getDataFromSelection);
            $('#get-files').click(getDataFromSelection);
        });
    };

    // Reads data from current document selection and displays a notification
    //function getDataFromSelection() {
    //    if (Office.context.document.getSelectedDataAsync) {
    //        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
    //            function (result) {
    //                if (result.status === Office.AsyncResultStatus.Succeeded) {
    //                    showNotification('The Excel files:', '"' + getFiles().xhr.status + '"');
    //                } else {
    //                    showNotification('Error:', result.error.message);
    //                }
    //            }
    //        );
    //    } else {
    //        app.showNotification('Error:', 'Reading selection data is not supported by this host application.');
    //    }
    //}

    function getDataFromSelection() {
        var xhr = new XMLHttpRequest();
        xhr.open("GET", "https://jsonplaceholder.typicode.com/posts/1", true);
 
        //xhr.open("GET", "https://author-revere.adobecqms.net/assets.json", true);
        //xhr.setRequestHeader("Authorization", "Basic " + btoa("towles:iy6azthsa6"));
        //xhr.withCredentials;

        xhr.send();
        showNotification('The Excel files:', '"' + xhr.responseURL + '"');
     }
    
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    function getFiles() {
        var xhr = new XMLHttpRequest();
        xhr.open("GET", "https://author-revere.adobecqms.net/api/assets.json", true);
        xhr.send();
        var myJson = xhr.responseText;
        console.log(xhr.status);
        console.log(xhr.statusText);
    };
})();