/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */
 
let speechFlag = false;
let popup=null;
let loader=true;
 
Office.onReady(async function (info) {
  displayStartingMessage("Hi! I'm Legal Agent, your personal Copilot agent. I can help you create documents, modify content, and insert useful information seamlessly. How can I assist you today?");
  let directLine1 = await initializeDirectLine();
if (info.host === Office.HostType.Word) {
 
document.getElementById("sendButton").onclick = async function () {
  const question = document.getElementById("userInput").value;
  if (question) {
    //document.getElementById("headerId").style.display = "none";
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
 
document.getElementById('startSpeechButton').addEventListener('click', function () {
  // Open a pop-up window to handle the speech
  if(speechFlag){
    mic.classList.toggle("recording");
    speechFlag=false;
    try{
      popup.close();
    }
    catch{
      console.log("popup closed")
    }
   
  }
  else{
  mic.classList.toggle("recording");
  popup = window.open(
    'speech.html',
    'SpeechRecognition',
    'width=1,height=1'
  );
  speechFlag = true;
  // Listen for messages from the pop-up window
  window.addEventListener("message", async function eventHandler(event){
      if (event.origin !== window.location.origin) return; // Security check
 
      // Get the recognized text from the pop-up
      transcript = event.data;
 
      // Insert recognized text into user input
      console.log(transcript);
      if(transcript==="NoSpeech"){
        popup.close();
        speechFlag=false;
        mic.classList.toggle("recording");
        window.removeEventListener("message", eventHandler);
      }
      else{
      document.getElementById("userInput").value = transcript;
      var question = document.getElementById("userInput").value  ;
    if (question) {
        document.getElementById("userInput").value ="";
        displayChatMessage(question, '', "User");
        const micButton=document.getElementById("startSpeechButton")//disable mic
        micButton.disabled=true;
        await getBotResponse(directLine1, question);
      }
      popup.close();
     
      mic.classList.toggle("recording");
      window.removeEventListener("message", eventHandler);
  }}, { once: true });
 
}
});
 
}
});
 
function displayStartingMessage(starter) {
  chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">${starter}</div>`;
}
 
 
// Display user question and bot response in chat window
async function displayChatMessage(question, response, role,directLine) {
  const chatWindow = document.getElementById("chatWindow");
  console.log("displayfunction called");
 
  // Check if response is valid and if attachments exist
  // eslint-disable-next-line no-constant-condition
  if (response && response.attachments && response.attachments.length > 0  ) {
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
            chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">${attachment.content.text}</div>`;
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
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">SOW content generated in document</div>`;
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText("S.O.W. content generated in document");
        });
       
        //speechFlag = false;  
        }
      }else if(response.speak==="Table"){
 
        insertResponseIntoDocumentAtCursor(response.text, "end");
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">Table has been generated in document</div>`;      
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText("Table has been generated in document");
        });
       
        //speechFlag = false;  
        }
      }
      else if(response.speak==="TableReplace"){
        let statusflag=false;
        statusflag=await insertResponseIntoDocumentAtCursor(response.text,"replace");
        console.log(statusflag);
        if(statusflag){
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">Table has been generated in document</div>`;      
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText("Table has been generated in document");
        });
       
        //speechFlag = false;  
        }}
        else{
          chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">No table is selected in the document</div>`;      
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText("No table is selected in the document");
        });
       
        //speechFlag = false;  
        }
        }
      }
      else if(response.speak==="Replace"){
        splitText=response.text
        textArray=splitText.split("|");
        replaceText(textArray[0],textArray[1]);
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">Replaced ${textArray[0]} with ${textArray[1]} </div>`;      
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText(`Replaced ${textArray[0]} with ${textArray[1]}`);
        });
       
        //speechFlag = false;  
        }
      }
      else if(response.speak==="Selected"){
        if(response.text==="Table"){
          console.log("fetching selected table")
          getSelectedTable(directLine);
        }
        else{
        console.log("fetching selected data")
        getSelectedText(directLine);  
        }
      }
      else if(response.speak==="paragraph"){
        setSelectedText(response.text);
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">Requested changes have been made in the document</div>`;  
        if(speechFlag){
          ensureVoicesLoaded(() => {
            speakText("Requested changes have been made in the document");
        });
       
        //speechFlag = false;  
        }
      }
      else if(response.speak==="interim"){
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">${response.text}</div>`;
        console.log("Calling loading function");
        displayLoading();
        if(speechFlag){
          ensureVoicesLoaded(async () => {
            await speakText(response.text);
            speechFlag=true;
        });
       
        }      
      }
      else if(response.speak==="interimSpeak"){
        console.log("Calling loading function");
        displayLoading();
        if(speechFlag){
          ensureVoicesLoaded(async () => {
            await speakText(response.text);
            speechFlag=true;
        });
       
        }      
      }
      else if(response.speak==="interimFinal"){
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">${response.text}</div>`;
        if(speechFlag){
          ensureVoicesLoaded(async () => {
            speakText(response.text);
           
        });
       
        }      
      }
   
      else if(response.text){
        chatWindow.innerHTML += `<div class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div class="message bot">${response.text}</div>`;
        document.getElementById("insertButton").style.display = "block";
        if(speechFlag){
          console.log("speaking bot message");
          ensureVoicesLoaded(() => {
            speakText(response.text);
        });
       
        //speechFlag = false;  
        }      
      }
     
    }
    else {
      if(question){
        document.getElementById("insertButton").style.display = "none";
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
 
async function insertResponseIntoDocumentAtCursor(response, insertAt) {
  if (insertAt === "end") {
    console.log("end of doc table");
    return await Word.run(async (context) => {
      const body = context.document.body;
      body.insertHtml(response, Word.InsertLocation.end);
      await context.sync();
      return true;
    });
  } else {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("parentTable");
      await context.sync();
 
      if (!selection.parentTable) {
          console.log("❌ No table selected.");
          return false;
      }
 
      const table = selection.parentTable;
      const tableRange = table.getRange(Word.RangeLocation.entire); // Capture range before deletion
      tableRange.load();
      await context.sync();
 
      console.log("Table found. Deleting...");
     
      // Insert a placeholder before deleting (to keep a valid reference)
      const placeholder = tableRange.insertText(" ", Word.InsertLocation.before);
      placeholder.load("text, address"); // Load placeholder info
      await context.sync();
 
      table.delete(); // Delete the table
      await context.sync();
      console.log("Table deleted.");
 
      console.log("Inserting new table...");
      placeholder.insertHtml(response, Word.InsertLocation.replace); // Insert new content
      await context.sync();
 
      console.log("New table inserted.");
      return true;
  });
  }}
 
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
      console.log("Testing activity on send: ", activity);
      console.log("Role", activity.from.role);
 
      if (activity.type === "message" && activity.from.id !== "10" && !activity.recipient) {
        console.log("Bot Response: ", activity.text);
        if(response.speak!=="Selected" ){
          try{
            document.getElementById("loader-container").remove();
            document.getElementById("loader").remove();
          }catch{
            console.log("could not find the element");
          }
          loader=true;
      }else{
        loader=false;
      }
        displayChatMessage(false, activity, activity.from.role,directLine);
      }
    });
    return directLine;
  } catch (error) {
    console.error("Error initializing DirectLine:", error);
  }
};
 
const getBotResponse = async function (directLine, question) {
  // call loading function
  console.log("Calling loading function");
  displayLoading();
  console.log("User:",question);
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
 
async function replaceText(oldText, newText) {
  await Word.run(async (context) => {
      let results = context.document.body.search(oldText, { matchCase: false });
      results.load("items");
      await context.sync();
      console.log("Results found:", results.items.length);
      for (let i = results.items.length - 1; i >= 0; i--) {
          let item = results.items[i];
          console.log("Replacing:", item.text);
          item.insertText(newText, Word.InsertLocation.replace);
      }
      await context.sync();
      console.log("ll instances replaced successfully.");
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
 
async function getSelectedTable(directLine) {
 
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("parentTable");
    await context.sync();
    if (selection.parentTable) {
        const table = selection.parentTable;
        table.load("values"); // Get table content as a 2D array
       
        await context.sync();
 
        const tableValues = table.values; // Array of rows with cell content
 
        let plainTextTable = "";
        tableValues.forEach(row => {
            plainTextTable += row.join(" | ") + "\n"; // Join cells with "|"
        });
 
        console.log(plainTextTable);
        await getBotResponse(directLine, plainTextTable);
         // Output the extracted table as plain text
    } else {
        console.log("No table selected by user.");
       
    }
});
 
 
}
 
async function setSelectedText(response) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(response, Word.InsertLocation.replace);
    await context.sync();
  });
}
 
async function speakText(text) {
  console.log("Testing Text to Speech");
 
  let voices = window.speechSynthesis.getVoices();
  console.log("Voices:", voices);
 
  let femaleVoice = voices.find(voice =>
    voice.name.includes("Female") ||
    voice.name.includes("Google UK English Female") ||
    voice.name.includes("Microsoft Zira") ||
    voice.name.includes("Samantha")
  );
 
  console.log("Set voice:", femaleVoice);
 
  const speech = new SpeechSynthesisUtterance(text);
 
  if (femaleVoice) {
    speech.voice = femaleVoice;
  } else {
    console.warn("Female voice not found. Using default voice.");
  }
 
  return new Promise((resolve) => {
    speech.onend = () => {
      console.log("Speech has finished.");
      speechFlag=false;
      // enable mic
      const micButton=document.getElementById("startSpeechButton")//disable mic
      micButton.disabled=false;
      console.log("mic enabled");
      resolve(true);
    };
 
    speech.onerror = (event) => {
      console.error("Speech error:", event.error);
      speechFlag=false;
      // enable mic
      const micButton=document.getElementById("startSpeechButton")//disable mic
      micButton.disabled=false;
      console.log("mic enabled");
      resolve(false);
    };
 
    window.speechSynthesis.speak(speech);
  });
}
 
 
// Load voices properly before calling the function
function ensureVoicesLoaded(callback) {
  let voices = window.speechSynthesis.getVoices();
 
  if (voices.length > 0) {
      callback();
  } else {
      window.speechSynthesis.onvoiceschanged = callback;
  }
}
 
 
function displayLoading(){
  if(loader){
    let loadingElement = document.getElementById("chatWindow");
    let loadingDots = [ ".", "..", "..."];
    let dotIndex = 1;
    loadingElement.innerHTML += `<div id="loader-container" class="bot-wrapper"><img width=20 height=20 src="assets/copilot.png"/> NoviPilot</div><div id = "loader" class="message bot">.</div>`;
    
    let loader = document.getElementById("loader");
    setInterval(() => {
      loader.innerText =loadingDots[dotIndex];
        dotIndex = (dotIndex + 1) % loadingDots.length;
    }, 500);
 }
}