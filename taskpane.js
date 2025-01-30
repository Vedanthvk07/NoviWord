/* eslint-disable no-undef */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    // Handle the Ask button click
    document.getElementById("askButton").onclick = async function () {
      const question = document.getElementById("userInput").value;
      if (question) {
        const response = await getChatbotResponse(question);
        displayChatMessage(question, response);
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
async function getChatbotResponse(question) {
  // For now, a simple mock response
  console.log("Testing");
  // Example Usage:
  // fetchGeminiResponse("Tell me a fun fact about space.");
  initializeDirectLine();
  return "This is a response to: " + question;
}

// Display user question and bot response in chat window
function displayChatMessage(question, response) {
  const chatWindow = document.getElementById("chatWindow");
  chatWindow.innerHTML += `<div class="user">You:</div><div>${question}</div>`;
  chatWindow.innerHTML += `<div class="bot">Bot:</div><div>${response}</div>`;
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
const initializeDirectLine = async function () {
  try {
    const response = await fetch(
      "https://148a369decc3eeda85b913c1e80b9a.da.environment.api.powerplatform.com/powervirtualagents/botsbyschema/cra27_agent123/directline/token?api-version=2022-03-01-preview"
    );
    const data = await response.json();
    console.log("Testing data token:" + JSON.stringify(data, null, 2));
    console.log("DirectLine Object:", window.DirectLine);
    const directLine = new window.DirectLine.DirectLine({ token: data.token });
    console.log("directLine:", directLine);
    if (!directLine || !directLine.activity$) {
      throw new Error("DirectLine instance failed to initialize");
    } else {
      console.log("Directline initialized");
    }

    directLine.activity$.subscribe((activity) => {
      console.log("Testing activity: ", activity);
      if (activity.type === "message" && activity.from.id !== "10" && !activity.recipient) {
        console.log("Testing response: ", activity.text);
      }
    });
  } catch (error) {
    console.error("Error initializing DirectLine:", error);
  }
};
console.log("Completed");

// Call the function

// async function fetchGeminiResponse(prompt) {
//   console.log("Testing inside gemini");
//   const apiKey = "AIzaSyB_ClqIjtTx2oL46vWfdKMFKUPB_YM3Ju8"; // Replace with your actual API key
//   const url = `https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key=${apiKey}`;

//   const requestBody = {
//     prompt: { text: prompt },
//     temperature: 0.7,
//   };

//   try {
//     const response = await fetch(url, {
//       method: "POST",
//       headers: {
//         "Content-Type": "application/json",
//       },
//       body: JSON.stringify(requestBody),
//     });

//     if (!response.ok) {
//       throw new Error(`HTTP error! Status: ${response.status}`);
//     }

//     const data = await response.json();
//     console.log("Gemini Response:", data);
//     return data;
//   } catch (error) {
//     console.error("Error fetching Gemini response:", error);
//   }
// }
