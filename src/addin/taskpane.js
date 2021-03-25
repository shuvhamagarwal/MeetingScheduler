// TEMPORARY CODE TO VERIFY ADD-IN LOADS
'use strict';

$(document).ready(function () {
  showWeekDays();
  $('#createMeeting').on('submit', function (e) { //use on if jQuery 1.7+
    createEvent(e);
  });
});

function showWeekDays() {
  $('input[type=radio][name=timeSlot]').change(function () {
    if (this.value == 'sameTime') {
      $('#weekDays').show();
      $('#weekDaysWithTime').hide();
    }
    else if (this.value == 'differentTime') {
      $('#weekDays').hide();
      $('#weekDaysWithTime').show();
    }
  });
}

Office.onReady(info => {
  // Only run if we're inside Excel
  if (info.host === Office.HostType.Outlook) {
    $(async function() {
      let apiToken = '';
      try {
        apiToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
        console.log(`API Token: ${apiToken}`);
      } catch (error) {
        console.log(`getAccessToken error: ${JSON.stringify(error)}`);
        // Fall back to interactive login
        showConsentUi();
      }

      // Call auth status API to see if we need to get consent
      const authStatusResponse = await fetch(`${getBaseUrl()}/auth/status`, {
        headers: {
          authorization: `Bearer ${apiToken}`
        }
      });

      const authStatus = await authStatusResponse.json();
      if (authStatus.status === 'consent_required') {
        showConsentUi();
      } else {
        // report error
        if (authStatus.status === 'error') {
          const error = JSON.stringify(authStatus.error,
            Object.getOwnPropertyNames(authStatus.error));
          showStatus(`Error checking auth status: ${error}`, true);
        } else {
          showMainUi();
        }
      }
    });
  }
});

// Handle to authentication pop dialog
let authDialog = undefined;

// Build a base URL from the current location
function getBaseUrl() {
  return location.protocol + '//' + location.hostname +
  (location.port ? ':' + location.port : '');
}

// Process the response back from the auth dialog
function processConsent(result) {
  const message = JSON.parse(result.message);

  authDialog.close();
  if (message.status === 'success') {
    showMainUi();
  } else {
    const error = JSON.stringify(message.result, Object.getOwnPropertyNames(message.result));
    showStatus(`An error was returned from the consent dialog: ${error}`, true);
  }
}

// Use the Office Dialog API to show the interactive
// login UI
function showConsentPopup() {
  const authDialogUrl = `${getBaseUrl()}/consent.html`;

  Office.context.ui.displayDialogAsync(authDialogUrl,
    {
      height: 60,
      width: 30,
      promptBeforeOpen: false
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        authDialog = result.value;
        authDialog.addEventHandler(Office.EventType.DialogMessageReceived, processConsent);
      } else {
        // Display error
        const error = JSON.stringify(error, Object.getOwnPropertyNames(error));
        showStatus(`Could not open consent prompt dialog: ${error}`, true);
      }
    });
}

// Inform the user we need to get their consent
function showConsentUi() {
  $('.authentication').empty();
  $('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: 'Consent for Microsoft Graph access needed'
  }).appendTo('.authentication');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'In order to access your calendar, we need to get your permission to access the Microsoft Graph.'
  }).appendTo('.authentication');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'We only need to do this once, unless you revoke your permission.'
  }).appendTo('.authentication');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'Please click or tap the button below to give permission (opens a popup window).'
  }).appendTo('.authentication');
  $('<button/>', {
    class: 'primary-button',
    text: 'Give permission'
  }).on('click', showConsentPopup)
  .appendTo('.authentication');
}

// Display a status
function showStatus(message, isError) {
  $('.status').empty();
  $('<div/>', {
    class: `status-card ms-depth-4 ${isError ? 'error-msg' : 'success-msg'}`
  }).append($('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: isError ? 'An error occurred' : 'Success'
  })).append($('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: message
  })).appendTo('.status');
}

function toggleOverlay(show) {
  $('.overlay').css('display', show ? 'block' : 'none');
}

function showMainUi() {
  $('.authentication').empty();
  $('.recurringMeetingContainer').show();
  console.log("done");
}

async function createEvent(evt) {
  evt.preventDefault();
  toggleOverlay(true);

  const apiToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
  const payload = {
    eventSubject: $('#title').val(),
    eventStart: $('#startDate').val(),
    eventEnd: $('#endDate').val(),
    formData: GetPayload()
  };
  console.log(payload);
  const requestUrl = `${getBaseUrl()}/graph/newevent`;

  const response = await fetch(requestUrl, {
    method: 'POST',
    headers: {
      authorization: `Bearer ${apiToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload)
  });

  if (response.ok) {
    showStatus('Event created', false);
  } else {
    const error = await response.json();
    showStatus(`Error creating event: ${JSON.stringify(error)}`, true);
  }

  toggleOverlay(false);
}

function GetPayload() {
  let title = $('#title').val();
  let startDate = $('#startDate').val();
  let endDate = $('#endDate').val();
  let timeSlotValue = $('input[name="timeSlot"]:checked').val();
  let isSameTime = false;
  if (timeSlotValue === "sameTime") {
    isSameTime = true;
  }
  var data = {};  
  data["isSameTime"]  = isSameTime;
  data["startDate"] = startDate;
  data["endDate"] = endDate;
  if (isSameTime) {
    var sameSlotArr = $('input[name="sameSlotDaysArr[]"]:checked');
    if (sameSlotArr.length > 0) {
      let startTime = $('#startTime1').val();
      let endTime = $('#endTime1').val();
      for(let i=0; i<sameSlotArr.length; i++) {
        let item = sameSlotArr[i];
        data[item.value] = {"startTime" : startTime, "endTime" : endTime};
      }
    } 
  } else {
    var differentSlotArr = $('input[name="differentSlotDays[]"]:checked');
      if (differentSlotArr.length > 0) {
        for(let i=0; i<differentSlotArr.length; i++) {
          let item = differentSlotArr[i];
          let itemsTr= $(item).parent().parent().parent();
          let secondtd = $(itemsTr).find("td:eq(1)");
          let startTime = secondtd.find("input[name='startTime']").val();
          let lastTd = $(itemsTr).find('td:last')
          let endTime = lastTd.find("input[name='endTime']").val();
          data[item.value] = {"startTime" : startTime, "endTime" : endTime};
        }
      }
  }
  return data;

}