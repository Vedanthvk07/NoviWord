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
