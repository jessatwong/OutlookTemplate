templateConfig=
{
    "Custom_Entities": {
        "OCCASIONS": [ "Birthday", "Funeral", "Get Well" ]
    },

    "Handoff": [
        {
            "Platform": "Web",
            "HandoffType": "GET",
            "URI": "http://www.bing.com/maps/default.aspx?where= ADDRESS "
        },
        {
            "Platform": "Web",
            "HandoffType": "POST",
            "URI": "https://posttestserver.com/post.php",
            "Address": "address"

        },
        {
            "Platform": "Windows",
            "HandoffType": "APP",
            "URI": "maps://?where= ADDRESS "
        },
        {
            "Platform": "iOs",
            "HandoffType": "APP",
            "URI": "http://maps.apple.com/?q= ADDRESS "
        },
        {
            "Platform": "Android",
            "HandoffType": "APP",
            "URI": "https://maps.google.com/maps?q= ADDRESS "
        }
    ],

    "Branding": {
        "Message": "Map addresses from your inbox!",
        "ButtonText": "Map",
        "ReturnMesssage": "Hope that was helpful"
    }

};