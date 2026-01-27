/* global Office, Word */

let initialized = false;
let exporting = false;

Office.onReady(() => {
    if (initialized) return;
    initialized = true;

    const btn = document.getElementById("exportBtn");
    btn.replaceWith(btn.cloneNode(true)); // remove previous listeners
    const newBtn = document.getElementById("exportBtn");

    newBtn.addEventListener("click", async () => {
        if (exporting) return; 
        exporting = true;

        // Disable button and grey it out
        newBtn.disabled = true;
        newBtn.style.opacity = 0.5;
        newBtn.style.cursor = "not-allowed";

        try {
            await exportSentences();
        } finally {
            // Re-enable button after requests finish
            exporting = false;
            newBtn.disabled = false;
            newBtn.style.opacity = 1;
            newBtn.style.cursor = "pointer";
        }
    });
});



async function exportSentences() {
    await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();

        const text = body.text;
        if (!text || !text.trim()) {
            alert("Document is empty.");
            return;
        }

        // Normalize text and split into sentences
        const sentences = text
            .replace(/\r?\n+/g, " ")
            .split(/(?<=[.!?])\s+(?=.)/)
            .map(s => s.trim())
            .filter(Boolean);

        // Optional: Download TXT
        // downloadTxt(sentences.join("\n"), "sentences.txt");

        // Call API for each sentence
        const results = [];
        for (const sentence of sentences) {
            try {
                const response = await fetch("http://127.0.0.1:5000/prompt", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ sentence })
                });

                if (!response.ok) {
                    console.error(`API error for sentence: ${sentence}`);
                    continue;
                }

                const data = await response.json();
                results.push(data);
            } catch (err) {
                console.error(`Request failed for sentence: ${sentence}`, err);
            }
        }

        // Display results in task pane
        displayResults(results);
    });
}

function downloadTxt(content, filename) {
    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();

    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// function displayResults(results) {
//     const container = document.getElementById("results");
//     container.innerHTML = ""; // Clear previous

//     results.forEach((item, idx) => {
//         const div = document.createElement("div");
//         div.style.border = "1px solid #ccc";
//         div.style.padding = "10px";
//         div.style.marginBottom = "8px";

//         div.innerHTML = `
//             <strong>Sentence ${idx + 1}:</strong><br>
//             <strong>Original:</strong> ${item.original.sentence}<br>
//             Lint Score: ${item.original.lint_score} | Valid: ${item.original.valid}<br>
//             <strong>Simplified:</strong> ${item.simplified.sentence}<br>
//             Lint Score: ${item.simplified.lint_score} | Valid: ${item.simplified.valid}<br>
//             <strong>Improvement:</strong> ${item.improvement}
//         `;
//         container.appendChild(div);
//     });
// }

let undoStack = []; // Stores previous suggestions for undo

function displayResults(results) {
    const container = document.getElementById("results");
    container.innerHTML = ""; // Clear previous

    results.forEach((item, idx) => {
        const div = document.createElement("div");
        div.className = "result-card";

        // Save original text for undo
        undoStack.push({
            original: item.original.sentence,
            simplified: item.simplified.sentence,
            index: idx
        });

        div.innerHTML = `
            <div>
                <strong>Original:</strong>
                <span class="original-text">${item.original.sentence}</span>
            </div>
            <div>
                Lint Score: ${item.original.lint_score} | Valid: ${item.original.valid}
            </div>
            <div>
                <strong>Simplified:</strong> 
                <span class="simplified-text">${item.simplified.sentence}</span>
            </div>
            <div>
                Lint Score: ${item.simplified.lint_score} | Valid: ${item.simplified.valid}
            </div>
            <div>
                <button class="accept-btn">Accepteren</button>
                <button class="modify-btn">Aanpassen</button>
                <button class="deny-btn">Weigeren</button>
            </div>
        `;

        // Button actions
        const acceptBtn = div.querySelector(".accept-btn");
        const modifyBtn = div.querySelector(".modify-btn");
        const denyBtn = div.querySelector(".deny-btn");
        const textSpan = div.querySelector(".simplified-text");

        acceptBtn.addEventListener("click", async () => {
            await replaceInWord(item.simplified.sentence, item.original.sentence);
            div.remove();
        });

        denyBtn.addEventListener("click", () => {
            div.remove(); // Remove suggestion
        });

        modifyBtn.addEventListener("click", () => {
            const input = document.createElement("input");
            input.type = "text";
            input.value = textSpan.textContent;
            input.style.width = "80%";

            const saveBtn = document.createElement("button");
            saveBtn.textContent = "Accepteren";
            const cancelBtn = document.createElement("button");
            cancelBtn.textContent = "Weigeren";

            const modifyDiv = document.createElement("div");
            modifyDiv.appendChild(input);
            modifyDiv.appendChild(saveBtn);
            modifyDiv.appendChild(cancelBtn);

            div.querySelector("div:last-child").replaceWith(modifyDiv);

            saveBtn.addEventListener("click", async () => {
                await replaceInWord(input.value, item.original.sentence);
                div.remove();
            });

            cancelBtn.addEventListener("click", () => {
                modifyDiv.replaceWith(div.querySelector("div:last-child"));
            });
        });

        container.appendChild(div);
    });

    // Undo button
    if (!document.getElementById("undoBtn")) {
        const undoBtn = document.createElement("button");
        undoBtn.id = "undoBtn";
        undoBtn.textContent = "Undo last changes";
        undoBtn.addEventListener("click", async () => {
            for (const entry of undoStack) {
                await replaceInWord(entry.original, entry.simplified);
            }
            undoStack = [];
            container.innerHTML = "";
        });
        container.parentElement.insertBefore(undoBtn, container);
    }
}

// Replace text in Word document
async function replaceInWord(newText, oldText) {
    await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();

        body.search(oldText).getFirst().insertText(newText, "Replace");
        await context.sync();
    });
}
