const predictBtn = document.getElementById("predictBtn");
const commentsEl = document.getElementById("comments");
const resultsBody = document.querySelector("#results tbody");

predictBtn.addEventListener("click", () => {
    const comments = commentsEl.value.split("\n").filter(c => c.trim() !== "");

    if (comments.length === 0) return;

    // Send to C#
    window.chrome.webview.postMessage(JSON.stringify({ comments }));
});

// Receive predictions from C#
window.chrome.webview.addEventListener("message", event => {
    const data = JSON.parse(event.data);

    resultsBody.innerHTML = "";
    for (let i = 0; i < data.comments.length; i++) {
        const row = document.createElement("tr");
        const cellComment = document.createElement("td");
        cellComment.textContent = data.comments[i];
        const cellPred = document.createElement("td");
        cellPred.textContent = data.predictions[i].toFixed(3);
        row.appendChild(cellComment);
        row.appendChild(cellPred);
        resultsBody.appendChild(row);
    }
});
