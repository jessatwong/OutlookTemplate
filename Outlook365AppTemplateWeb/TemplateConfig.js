templateConfig=
{
    "Custom_Entities": {
        "OCCASIONS": [ "Birthday", "Funeral", "Get Well" ]
    },

    "Handoff": [
        {
            "Platform": "Web",
            "HandoffType": "GET",
            "URI": "http://www.bing.com/maps/default.aspx?where= ADDRESS ",
            "ButtonValue": "Map Using Web (GET)"
            },
        {
            "Platform": "Web",
            "HandoffType": "POST",
            "URI": "https://posttestserver.com/post.php",
            "Address": "address",
            "ButtonValue": "Map Using Web (POST)"

        },
        {
            "Platform": "Windows",
            "HandoffType": "APP",
            "URI": "maps://?where= ADDRESS ",
            "ButtonValue": "Map Using Windows (APP)"
        },
        {
            "Platform": "iOs",
            "HandoffType": "APP",
            "URI": "http://maps.apple.com/?q= ADDRESS ",
            "ButtonValue": "Map Using iOS (APP)"
        },
        {
            "Platform": "Android",
            "HandoffType": "APP",
            "URI": "https://maps.google.com/maps?q= ADDRESS ",
            "ButtonValue": "Map Using Android (APP)"
        }
    ],

    "Branding": {
        "Message": "Map addresses from your inbox!",
        "ButtonText": "Map",
        "ReturnMesssage": "Hope that was helpful"
    }

};