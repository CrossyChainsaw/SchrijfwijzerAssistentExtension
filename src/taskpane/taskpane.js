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
    container.innerHTML = ""; 

    const improvedResults = results.filter(item => 
        item.simplified.lint_score <= item.original.lint_score
    );

    if (improvedResults.length === 0) {
        container.innerHTML = "<em>No suggestions improved the lint score.</em>";
        return;
    }

    // --- 1. Create the Undo Button ONCE ---
    if (!document.getElementById("undoBtn")) {
        const undoBtn = document.createElement("button");
        undoBtn.id = "undoBtn";
        undoBtn.textContent = "Undo last change";
        undoBtn.style.marginBottom = "10px"; // Give it some space
        
        undoBtn.addEventListener("click", async () => {
            if (undoStack.length === 0) return;

            const lastChange = undoStack.pop();
            await replaceInWord(lastChange.original, lastChange.applied);
            
            updateUndoButtonState(); 
        });

        // Insert at the top of the container's parent
        container.parentElement.insertBefore(undoBtn, container);
    }
    
    // Set the initial state (disabled because stack is empty at start)
    updateUndoButtonState();

    improvedResults.forEach((item, idx) => {
        const div = document.createElement("div");
        div.className = "result-card";

        // NOTE: I removed the undoStack.push that was here. 
        // We only push when the user clicks "Accepteren".

        div.innerHTML = `
            <div><strong>Original:</strong> <span class="original-text">${item.original.sentence}</span></div>
            <div>Lint Score: ${item.original.lint_score}</div>
            <div><strong>Simplified:</strong> <span class="simplified-text">${item.simplified.sentence}</span></div>
            <div>Lint Score: ${item.simplified.lint_score}</div>
            <div class="button-group">
                <button class="accept-btn">Accepteren</button>
                <button class="modify-btn">Aanpassen</button>
                <button class="deny-btn">Weigeren</button>
            </div>
        `;

        const acceptBtn = div.querySelector(".accept-btn");
        const modifyBtn = div.querySelector(".modify-btn");
        const denyBtn = div.querySelector(".deny-btn");

        // --- 2. Handle Accept ---
        acceptBtn.addEventListener("click", async () => {
            await replaceInWord(item.simplified.sentence, item.original.sentence);

            undoStack.push({
                original: item.original.sentence,
                applied: item.simplified.sentence
            });

            updateUndoButtonState(); 
            div.remove();
        });

        denyBtn.addEventListener("click", () => div.remove());

        // --- 3. Handle Modify ---
        modifyBtn.addEventListener("click", () => {
            const originalButtonsDiv = div.querySelector(".button-group");
            originalButtonsDiv.style.display = "none";

            const modifyDiv = document.createElement("div");
            const input = document.createElement("input");
            input.type = "text";
            input.value = item.simplified.sentence;
            input.style.width = "70%";

            const saveBtn = document.createElement("button");
            saveBtn.textContent = "Opslaan";
            
            const cancelBtn = document.createElement("button");
            cancelBtn.textContent = "Annuleren";

            modifyDiv.append(input, saveBtn, cancelBtn);
            div.appendChild(modifyDiv);

            saveBtn.addEventListener("click", async () => {
                await replaceInWord(input.value, item.original.sentence);
                
                undoStack.push({
                    original: item.original.sentence,
                    applied: input.value
                });

                updateUndoButtonState();
                div.remove();
            });

            cancelBtn.addEventListener("click", () => {
                modifyDiv.remove();
                originalButtonsDiv.style.display = "block";
            });
        });

        container.appendChild(div);
    });
}


// Replace text in Word document
async function replaceInWord(newText, oldText) {
    await Word.run(async (context) => {
        const results = context.document.body.search(oldText);
        results.load("items");
        await context.sync();

        if (results.items.length > 0) {
            // Grab the first instance found
            results.items[0].insertText(newText, Word.InsertLocation.replace);
            await context.sync();
        } else {
            console.warn("Could not find text to replace:", oldText);
        }
    });
}

// Function to toggle the button's appearance and functionality
function updateUndoButtonState() {
    const undoBtn = document.getElementById("undoBtn");
    if (!undoBtn) return;

    if (undoStack.length === 0) {
        undoBtn.disabled = true;
        undoBtn.style.opacity = 0.5;
        undoBtn.style.cursor = "not-allowed";
    } else {
        undoBtn.disabled = false;
        undoBtn.style.opacity = 1;
        undoBtn.style.cursor = "pointer";
    }
}

// Update your Undo Button creation inside displayResults
if (!document.getElementById("undoBtn")) {
    const undoBtn = document.createElement("button");
    undoBtn.id = "undoBtn";
    undoBtn.textContent = "Undo last change";
    
    undoBtn.addEventListener("click", async () => {
        if (undoStack.length === 0) return;

        const lastChange = undoStack.pop();
        await replaceInWord(lastChange.original, lastChange.applied);
        
        // Refresh button state after popping
        updateUndoButtonState(); 
    });

    container.parentElement.insertBefore(undoBtn, container);
    updateUndoButtonState(); // Set initial disabled state
}
