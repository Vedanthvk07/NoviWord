/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */

//const { split } = require("core-js/fn/symbol");

Office.onReady(async function (info) {
  displayStartingMessage("Hi, I am your word assistant bot-NoviWord");
  let directLine1 = await initializeDirectLine();
if (info.host === Office.HostType.Word) {
  //let flag=true;
  
document.getElementById("askButton").onclick = async function () {
  const question = document.getElementById("userInput").value;
  if (question) {
    document.getElementById("headerId").style.display = "none";
    displayChatMessage(question, '', "User",directLine1);
      await getBotResponse(directLine1, question);
    

  }
};

document.getElementById("userInput").addEventListener("keydown", async function (event) {
  if (event.key === "Enter") {
    // Check if Enter key is pressed
    event.preventDefault(); // Prevents the default behavior (like submitting a form)

    const question = document.getElementById("userInput").value;
    if (question) {
      //document.getElementById("headerId").style.display = "none";
        displayChatMessage(question, '', "User",directLine1);
      await getBotResponse(directLine1, question);
      
  }
}});

// Handle the Insert button click
document.getElementById("insertButton").onclick = async function () {
  const response = document.getElementById("chatWindow").lastChild
    ? document.getElementById("chatWindow").lastChild.innerText
    : "";
  if (response) {
    await insertResponseIntoDocument(response);
  }
};
}
});

function displayStartingMessage(starter) {
  const chatWindow = document.getElementById("chatWindow");
  chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">${starter}</div>`; 
  // getDocProperties();    
}

// async function getDocProperties() {
//   await Word.run(async (context) => {
//   let docProperties = context.document.properties;
//     docProperties.load("title");
    
//     await context.sync();
//     console.log("Document Name: ", docProperties.title);
//     console.log("Document props: ", docProperties);
//   });
// }
// Display user question and bot response in chat window
function displayChatMessage(question, response, role,directLine) {
  const chatWindow = document.getElementById("chatWindow");

  // Check if response is valid and if attachments exist
  // eslint-disable-next-line no-constant-condition
  if (response && response.attachments && response.attachments.length > 0 && false) {
    response.attachments.forEach((attachment) => {
      // Check if attachment content has 'buttons' and 'signin' type
      if (attachment.content && attachment.content.buttons && attachment.content.buttons.length > 0) {
        attachment.content.buttons.forEach((button) => {
          if (button.type === "signin") {
            // Create a sign-in button
            const signinButton = document.createElement("button");
            signinButton.innerText = button.title || "Sign In"; // Default title to "Sign In"
            signinButton.classList.add("ms-Button", "ms-Button--primary");

            // Open the sign-in URL when the button is clicked
            signinButton.onclick = () => {
              window.open(button.value, "_blank"); // Open the sign-in URL in a new tab
            };

            // Display the bot's message
            chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">${attachment.content.text}</div>`;
            chatWindow.appendChild(signinButton); // Add the button after the message
          }
        });
      }
    });
  } else {
    // Regular message display if no attachments
    if (role === "bot") {
      if(response.speak==="Generate"){

        insertResponseIntoDocument(response.text);
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">SOW content generated in document</div>`; 
      
      }else if(response.speak==="Table"){

        insertResponseIntoDocumentAtCursor(response.text);
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">Table generated in document</div>`;      
      
      }
      else if(response.speak==="Replace"){
        splitText=response.text
        textArray=splitText.split("|");
        replaceText(textArray[0],textArray[1]);
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">Replaced ${textArray[0]} with ${textArray[1]} </div>`;      
      }else if(response.speak==="Selected"){
        getSelectedText(directLine);  
      }
    
      else if(response.text){
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviWord</div><div class="message bot">${response.text}</div>`;      
      }
      
    } 
    else {
      if(question){
      
        chatWindow.innerHTML += `<div class="user-wrapper">You</div><div class="message user">${question}</div>`;      }
      
    }
  }
  scrollToBottom();
  // Clear the input field
  document.getElementById("userInput").value = "";
}

// Function to insert the response into the Word document
async function insertResponseIntoDocument(response) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertHtml(response, Word.InsertLocation.end);
    await context.sync();
  });
}

async function insertResponseIntoDocumentAtCursor(response) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertHtml(response, Word.InsertLocation.after);
    await context.sync();
  });
}
const initializeDirectLine = async function () {
  try {
    const response = await fetch(
      "https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview"
    );
    const data = await response.json();
    
    const directLine = new window.DirectLine.DirectLine({ token: data.token });
    

    if (!directLine || !directLine.activity$) {
      throw new Error("DirectLine instance failed to initialize");
    }

    directLine
      .postActivity({
        from: { id: "10", name: "User" },
        type: "message",
        text: "Hi",
      })
      .subscribe(
        (id) => console.log("Message sent with ID:", id),
        (error) => console.error("Error sending message:", error)
      );

    directLine.activity$.subscribe((activity) => {
      console.log("Testing activity: ", activity);
      console.log("Role", activity.from.role);
      if (activity.type === "message" && activity.from.id !== "10" && !activity.recipient) {
        console.log("Bot Response: ", activity.text);
        displayChatMessage(false, activity, activity.from.role,directLine);
        
      }
    });
    return directLine;
  } catch (error) {
    console.error("Error initializing DirectLine:", error);
  }
};

const getBotResponse = async function (directLine, question) {
  directLine
    .postActivity({
      from: { id: "10", name: "User" },
      type: "message",
      text: question,
    })
    .subscribe(             // calls the subscription already created in InitializeDirectline method
      (id) => console.log("Message sent with ID:", id),
      (error) => console.error("Error sending message:", error)
    );
  
}
function scrollToBottom() {
  const chatWindow = document.getElementById("chatWindow");
  setTimeout(() => {
    chatWindow.scrollTop = chatWindow.scrollHeight;
  }, 100); // Timeout ensures scroll happens after the new message is rendered
}

async function replaceText(oldText,NewText){
  await Word.run(async (context) => {
    let results = context.document.body.search(oldText);
    results.load();
    await context.sync();
    
    results.items.forEach(item => {
        item.insertText(NewText, Word.InsertLocation.replace);
    });
    
    await context.sync();
});
}

async function getSelectedText(directLine) {
  await Word.run(async (context) => {
    let range = context.document.getSelection();
    range.load("text");
    await context.sync();
    SelText=range.text;
    await getBotResponse(directLine, SelText);
});
  
}

let recognition = null;
 
document.getElementById("speakButton").addEventListener("click", async () => {
 
  try {
    await navigator.mediaDevices.getUserMedia({ audio: true });
    document.getElementById("output").innerText = "Listening...";
    console.log("Microphone access granted.");
    startVoiceInput();
  } catch (error) {
    console.error("Microphone access denied:", error);
  }
  // Request permission from a user to access their camera and microphone.
if (Office.context.platform === Office.PlatformType.OfficeOnline) {
  const deviceCapabilities = [
      Office.DevicePermissionType.camera,
      Office.DevicePermissionType.microphone
  ];
  Office.devicePermission
      .requestPermissions(deviceCapabilities)
      .then((isGranted) => {
          if (isGranted) {
              console.log("Permission granted.");
              // Reload your add-in before you run code that uses the device capabilities.
              location.reload();
          } else {
              console.log("Permission has been previously granted and is already set in the iframe.");

              // Since permission has been previously granted, you don't need to reload your add-in.

              // Do something with the device capabilities.
          }
      })
      .catch((error) => {
          console.log("Permission denied.");
          console.error(error);

          // Do something when permission is denied.
      });
}
});
 
function startVoiceInput() {
  // If a recognition instance already exists, stop it before starting a new one
  if (recognition && recognition.abort) {
    recognition.abort(); // This will stop the ongoing recognition session
  }
 
  recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
  recognition.lang = "en-US";
  recognition.interimResults = false;
  recognition.maxAlternatives = 1;
  recognition.continuous = true;
 
  // recognition.start();
   // trying continuos
  recognition.onstart = function () {
    console.log("Speech recognition started.");
  };
   // trying continuos
  recognition.onresult = function (event) {
    const transcript = event.results[0][0].transcript;
    console.log("Recognized text:", transcript);
    // insertTextIntoWord(transcript);
    document.getElementById("userInput").value = transcript;
    document.getElementById("output").innerText = transcript
  };
 
  recognition.onerror = function (event) {
    console.error("Speech recognition error:", event.error);
     // trying continuos
    if (event.error === 'aborted') {
      console.log("Restarting recognition...");
      recognition.start();
    }
     // trying continuos
  };
  // trying continuos
  recognition.onend = function () {
    console.log("Speech recognition ended.");
  };
 
  recognition.start();
  //  // trying continuos
}
