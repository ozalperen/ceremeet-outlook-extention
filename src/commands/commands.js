//create meeting id
function makeid(length) {
    var result           = '';
    var characters       = '0123456789';
    var charactersLength = characters.length;
    for ( var i = 0; i < length; i++ ) {
      result += characters.charAt(Math.floor(Math.random() * 
 charactersLength));
   }
   return result;
};

var meeting_id = makeid(12);

//create meeting pass
function makepass(length) {
    var result           = '';
    var characters       = 'abcdefghijklmnoprstuvwxyz';
    var charactersLength = characters.length;
    for ( var i = 0; i < length; i++ ) {
      result += characters.charAt(Math.floor(Math.random() * 
 charactersLength));
   }
   return result;
};

let meeting_pass = makepass(15);

//create meeting folder
    var mkdir_req = "https://eldercare.unknownland.org/file_mkdir"; //$.URLEncode(name);

    $.ajax({
        type: 'POST',
        url: mkdir_req,
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        data: JSON.stringify( { "meeting_id": meeting_id, "meeting_pass": meeting_pass} ),
        success : function(data) {
          //Success block  
        },
       error: function (xhr,ajaxOptions,throwError){
        //Error block 
      },
    });


//Write to body
const newBody = '<br>' + '<hr>' + '<br>' +
    '<h2>Ceremeet toplantı daveti</h2>' +
    '<strong>Ceremeet uygulaması üzerinden toplantıya katılın.</strong>' +
    '<br><br>' +
    '<a href="ceremeet://ceremeet.com/' + meeting_id + '?pwd=' +  meeting_pass + '" target="_blank">Ceremeet Giriş</a>' +
    '<br><br>' +
    '<strong>veya toplantı bilgileri ile giriş yapın:</strong>' +
    '<br><br>' +
    'Toplantı Odası: ' + meeting_id +
    '<br><br>' +
    'Toplantı Şifresı: ' + meeting_pass +
    '<br><br>' +
    '<a href="https://nextcloud.unknownland.org/s/ceremeet?path=' + meeting_id + '" target="_blank">Sunum Ekle</a>' +
    '<br><br>' +
    '<a href="https://nextcloud.unknownland.org/s/aK9Q8Ed4D3Mwd5Z/download"' + '" target="_blank">Ceremeet Uygulamasını İndirin.</a>' +
    '<br><br>';
    
let mailboxItem;


// Office is ready.
Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

// 2. How to define and register a function command named `insertCeremeetMeeting` (referenced in the manifest)
//    to update the meeting body with the online meeting details.
function insertCeremeetMeeting(event) {
    // Get HTML body from the client.
    mailboxItem.body.getAsync("html",
        { asyncContext: event },
        function (getBodyResult) {
            if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                updateBody(getBodyResult.asyncContext, getBodyResult.value);
            } else {
                console.error("Failed to get HTML body.");
                getBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );
}
// Register the function.
Office.actions.associate("insertCeremeetMeeting", insertCeremeetMeeting);

// 3. How to implement a supporting function `updateBody`
//    that appends the online meeting details to the current body of the meeting.
function updateBody(event, existingBody) {
    // Append new body to the existing body.
    mailboxItem.body.setAsync(existingBody + newBody,
        { asyncContext: event, coercionType: "html" },
        function (setBodyResult) {
            if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                setBodyResult.asyncContext.completed({ allowEvent: true });
            } else {
                console.error("Failed to set HTML body.");
                setBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );
}