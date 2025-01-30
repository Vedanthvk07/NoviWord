/* eslint-disable no-undef */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    // Handle the Ask button click
    document.getElementById("askButton").onclick = async function () {
      const question = document.getElementById("userInput").value;
      if (question) {
        //const response =
        //await
        initializeDirectLine(question);
        // displayChatMessage(question, response);
      }
    };

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

// Function to get the chatbot's response (simple hardcoded response or integrate with an API)
// async function getChatbotResponse(question) {
//   // For now, a simple mock response
//   // console.log("Testing");
//   // Example Usage:
//   // fetchGeminiResponse("Tell me a fun fact about space.");
//   initializeDirectLine(question);
//   // return "This is a response to: " + question;
// }

// Display user question and bot response in chat window
function displayChatMessage(question, response, role) {
  const chatWindow = document.getElementById("chatWindow");
  if (role === "bot") {
    chatWindow.innerHTML += `<div class="bot">Bot:</div><div>${response}</div>`;
  } else {
    chatWindow.innerHTML += `<div class="user">You:</div><div>${question}</div>`;
  }

  document.getElementById("userInput").value = ""; // Clear input field
}

// Function to insert the response into the Word document
async function insertResponseIntoDocument(response) {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.insertText(response, Word.InsertLocation.end);
    await context.sync();
  });
}
const initializeDirectLine = async function (question) {
  try {
    const response = await fetch(
      "https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview"
    );
    const data = await response.json();
    // console.log("Testing data token:" + JSON.stringify(data, null, 2));
    // console.log("DirectLine Object:", window.DirectLine);
    const directLine = new window.DirectLine.DirectLine({ token: data.token });
    // console.log("directLine*******", directLine);
    // console.log("DirectLine instance:", directLine);
    // console.log("DirectLine activity$:", directLine.activity$);

    if (!directLine || !directLine.activity$) {
      throw new Error("DirectLine instance failed to initialize");
    }
    directLine
      .postActivity({
        from: { id: "10", name: "User" },
        type: "message",
        text: question,
      })
      .subscribe(
        (id) => console.log("Message sent with ID:", id),
        (error) => console.error("Error sending message:", error)
      );

    directLine.activity$.subscribe((activity) => {
      console.log("Testing activity: ", activity);
      console.log("Role*******", activity.from.role);
      if (activity.type === "message" && activity.from.id !== "10" && !activity.recipient) {
        console.log("Testing response: ", activity.text);
        displayChatMessage(question, activity.text, activity.from.role);
      }
    });
  } catch (error) {
    console.error("Error initializing DirectLine:", error);
  }
};
