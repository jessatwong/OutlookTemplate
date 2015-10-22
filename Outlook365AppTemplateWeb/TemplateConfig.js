templateConfig=
{
    "Custom_Entities": {
        "OCCASIONS": [ "Birthday", "Funeral", "Get Well" ]
    },

    /*
        @HandoffType: must be either POST or GET
        @Full EntityTypes:
            ["ADDRESS", "CONTACT_NAME", "CONTACT_PHONE", "CONTACT_EMAIL", "MEETING_NAME", "TASK_NAME"]
        @InWindow: true or false to dictate if the link will open a new tab
        @URI: Must include single space before and after variable to replace in caps, i.e " ADDRESS "
        @ButtonValue: Text value on button
    */
    "Handoff": [
        {
            "Platform": "Web",
            "HandoffType": "GET",
            "EntityTypes": ["ADDRESS"],
            "InWindow": true,
            "URI": "https://www.bing.com/mapspreview?where= ADDRESS ",
            "ButtonValue": "Map Using Web (GET)"
            },
        {
            "Platform": "Web",
            "HandoffType": "POST",
            "EntityTypes": ["ADDRESS"],
            "InWindow": false,
            "URI": "https://posttestserver.com/post.php",
            "PostKey": "Address",
            "ButtonValue": "Map Using Web (POST)"

        },
        {
            "Platform": "Windows",
            "HandoffType": "APP",
            "EntityTypes": ["ADDRESS"],
            "InWindow": false,
            "URI": "maps://?where= ADDRESS ",
            "ButtonValue": "Map Using Windows (APP)"
        },
        {
            "Platform": "iOs",
            "HandoffType": "APP",
            "EntityTypes": ["ADDRESS"],
            "InWindow": false,
            "URI": "https://maps.apple.com/?q= ADDRESS ",
            "ButtonValue": "Map Using iOS (APP)"
        },
        {
            "Platform": "Android",
            "HandoffType": "APP",
            "EntityTypes": ["ADDRESS"],
            "InWindow": false,
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