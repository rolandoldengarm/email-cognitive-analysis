/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayItemDetails();
        });
    };

    function getBody() {
        var _item = Office.context.mailbox.item;
        var body = _item.body;

        // Get the body asynchronous as text
        body.getAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
                console.log("Error retrieving body");
            }
            else {
                // Show data

                console.log('Body', asyncResult.value.trim());
                processBody(asyncResult.value.trim());
            }
        });
    }

    function processBody(body) {
        console.log("processing body");
        var url = "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment";
        $.ajax({
            url: url,
            type: 'post',
            headers: {
                'Ocp-Apim-Subscription-Key': '8b8553844204419da281d9532267513a'
            },
            contentType: "application/json",
            data: JSON.stringify({
                "documents": [
                  {
                      "id": "body",
                      "text": body
                  }
                ]
            }),
            success: function (data) {
                console.info(data);
                var score = data.documents[0].score;
                var result = "This appears to be a positive message!";
                var color = "green";
                if (score < 0.5) {
                    result = "This appears to be a negative message";
                    color = "red";
                }
                $("#results").html(result);
                $("#results").css('color', color);
            },
            error: function (err) {
                console.error(err);
            }
        });

    }    
    function displayItemDetails() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);        
        getBody();       

        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            from = Office.cast.item.toMessageRead(item).from;
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            from = Office.cast.item.toAppointmentRead(item).organizer;
        }
    }
})();