/// <reference path="/Scripts/FabricUI/MessageBanner.js" />
/// <reference path="App.js" />
// global app

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
            
            // If not using Excel 2016, use fallback logic.
            //if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
            //    $('#loginbuttontext').text("Login to CRM");
            //    return;
            //}

 
            //Event Handlers
            $('#loginbuttontext').text("Login to CRM");
            $('#loginbutton').click(loginToCRM);
        });
    }
    
    function loginToCRM() {
        //check to see if there is an OAuth Token cached in the cookie
        //app.addinName will need to be different for every Add-In
        var token = app.getCookie(app.addinName)
        if (!token) {
            var tokenParams = {};
            tokenParams.authServer = 'https://login.windows.net/common/oauth2/authorize?';
            tokenParams.responseType = 'token';
            tokenParams.replyUrl = location.href.split("?")[0];

            //THESE tokenParams need to be changed for your application
            //tokenParams.clientId = 'Your-app-clientId-goes-here';
            tokenParams.clientId = 'b3beda96-6b4e-4200-857f-cf75c5425054';
            tokenParams.resource = "https://programmerbiz.github.io/index.html";
            //tokenParams.resource = "https://yoursite.sharepoint.com";

            app.getToken(tokenParams);
        } else {
            //we have a token therefore carry on
            app.tokenCallback(token)
        }
    };


    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
