/// <reference path="../App.js" />




(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            //inject content
            injectContent();

            //inject handoff buttons
            injectHandoff();
        });
    };
   
    function isHandoffType(handoffType) {
        return templateConfig.Handoff.HandoffTypes.indexOf(handoffType) > -1;
    }


    function injectContent(){
        $("#messageContainer").text(templateConfig.Branding.Message);
    }

    function injectHandoff() {
        // Extract entites and format handoff strings
        var item = Office.context.mailbox.item;
        var entities = item.getEntities();
        //Format Address
        //if  Address is in any of the handoff strings (GET, POST and APP URI) format the value in the string (TBI find a way to abstract this)
        if ((isHandoffType("GET") > -1) & (templateConfig.Handoff.GETString.indexOf(" ADDRESS ") > -1)) templateConfig.Handoff.GETString = templateConfig.Handoff.GETString.replace(" ADDRESS ", entities.addresses[0]);
        if ((isHandoffType("POST") > -1) & (templateConfig.Handoff.POSTString.indexOf(" ADDRESS ") > -1)) templateConfig.Handoff.POSTString = templateConfig.Handoff.POSTString.replace(" ADDRESS ", entities.addresses[0]);
        if ((isHandoffType("APP") > -1) & (templateConfig.Handoff.APPUri.indexOf(" ADDRESS ") > -1)) templateConfig.Handoff.APPUri = templateConfig.Handoff.APPUri.replace(" ADDRESS ", entities.addresses[0]);
        console.log(templateConfig.Handoff.APPUri);

        //TBI Other Entites from https://msdn.microsoft.com/en-us/library/office/fp161071.aspx 
        
        //TBI Custom Entities         

    
      //Inject the Handoff Forms
      //   Check if there is a GET handoff in HandoffTypes
        if (isHandoffType("GET")) {
            // Generate a hyperlink button with an href  of the GETString
            var getlink = $("<a href=\"" + templateConfig.Handoff.GETString + "\" target=\"_blank\">"+ templateConfig.Branding.ButtonText + " Using Web</a><br> ");
            var getlinkWStyle = $("<br><div class='ms-Button'  id='get-data-from-selection'> <span class='ms-Button-icon'><i class='ms-Icon ms-Icon--plus'></i></span> <a href=\"" + templateConfig.Handoff.GETString + "\" target=\"_blank\">" + templateConfig.Branding.ButtonText + " Using Web</a> </div>");

            //Grab the buttonContainer div from Home.html and inject the new hyperlink
            $("#buttonContainer").append(getlinkWStyle);
        }
            
      // Check if there is a APP handoff in HandoffTypes
        if (isHandoffType("APP")) {
            // Generate a hyperlink button with an href  of the GETString
            var applink =$("<a href=\"" + templateConfig.Handoff.APPUri + "\" target=\"_blank\">" + templateConfig.Branding.ButtonText + " Using App</a><br> ");

            var applinkWStyle = $("<div class='ms-Button'  id='get-data-from-selection'> <span class='ms-Button-icon'><i class='ms-Icon ms-Icon--plus'></i></span> <a href=\"" + templateConfig.Handoff.APPUri + "\" target=\"_blank\">" + templateConfig.Branding.ButtonText + " Using App</a> </div>");
            
            //Grab the buttonContainer div from Home.html and inject the new hyperlink
            $("#buttonContainer").append(applinkWStyle);
        }

        if (isHandoffType("POST")) {

        }
        /* Step 3 (TBI) Check if there is a POST handoff in HandoffTypes
            Step A: Generate a form with Method = POST, target _blank and action = postString
            Step B: Generate a button that submits the form
            Step C: Grab the buttonContainer divs from Home.html and inject the new post form and submit button 

      */
        
    }

})();