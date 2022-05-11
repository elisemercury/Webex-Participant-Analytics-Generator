/*!
* Start Bootstrap - Landing Page v6.0.5 (https://startbootstrap.com/theme/landing-page)
* Copyright 2013-2022 Start Bootstrap
* Licensed under MIT (https://github.com/StartBootstrap/startbootstrap-landing-page/blob/master/LICENSE)
*/
// This file is intentionally blank
// Use this file to add JavaScript to your project


function onMeetingNrSubmitted(){
    var meeting_nr = document.getElementById("meetingNr").value;
    console.log(meeting_nr);
    var timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    console.log(timezone);
    var params = {};
    params["meeting_nr"] = meeting_nr;
    params["timezone"] = timezone;
    const data = JSON.stringify(params);
    const current_url = window.location.href;
    console.log(current_url);
    console.log(data);
    var xhr = new XMLHttpRequest();
    xhr.open("POST", current_url, true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.send(data);
}