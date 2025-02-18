let task = null;  // Stores the detected task

function sendMessage() {
    let inputField = document.getElementById("user-input");
    let message = inputField.value.trim();
    
    if (message === "") return;

    appendMessage(message, "user-message");

    fetch("/chat", {
        method: "POST",
        body: JSON.stringify({ message }),
        headers: { "Content-Type": "application/json" }
    })
    .then(response => response.json())
    .then(data => {
        task = data.task;  // Store detected task
        appendMessage(data.response, "bot-message");
    });

    inputField.value = "";
}

function uploadFile() {
    let fileInput = document.getElementById("file-upload");
    let file = fileInput.files[0];

    if (!file || !task) {
        appendMessage("Please describe your request before uploading a file.", "bot-message");
        return;
    }

    let formData = new FormData();
    formData.append("file", file);
    formData.append("task", task);

    appendMessage("Uploading file: " + file.name, "user-message");

    fetch("/process-file", {
        method: "POST",
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        appendMessage(data.response, "bot-message");
        if (data.file_url) {
            appendMessage(`Download your processed file: <a href="${data.file_url}" target="_blank">Click here</a>`, "bot-message");
        }
    });
}

function appendMessage(text, className) {
    let chatBox = document.getElementById("chat-box");
    let messageDiv = document.createElement("div");
    messageDiv.innerHTML = text;
    messageDiv.classList.add(className);
    chatBox.appendChild(messageDiv);
    chatBox.scrollTop = chatBox.scrollHeight;
}

function handleKeyPress(event) {
    if (event.key === "Enter") {
        sendMessage();
    }
}
