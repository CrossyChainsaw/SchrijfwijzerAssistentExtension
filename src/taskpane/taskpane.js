/* global Office, Word */

// Customization
let improvementMethod;
improvementMethod = "better"; // only if simplified is better than original
// improvementMethod = "b1"; // only if simplified is between B1 range (36.18-50.07)




let initialized = false;
let exporting = false;
let undoStack = []; // Stores { fullItem, appliedText }

Office.onReady(() => {
    if (initialized) return;
    initialized = true;

    const btn = document.getElementById("exportBtn");
    btn.replaceWith(btn.cloneNode(true));
    const newBtn = document.getElementById("exportBtn");

    newBtn.addEventListener("click", async () => {
        if (exporting) return;
        exporting = true;

        newBtn.disabled = true;
        newBtn.style.opacity = 0.5;
        newBtn.style.cursor = "not-allowed";

        try {
            await exportSentences();
        } finally {
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

        const sentences = text
            .replace(/\r?\n+/g, " ")
            .split(/(?<=[.!?])\s+(?=.)/)
            .map(s => s.trim())
            .filter(Boolean);

        const results = [];
        for (const sentence of sentences) {
            try {
                const response = await fetch("http://127.0.0.1:5000/prompt", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ sentence })
                });

                if (response.ok) {
                    const data = await response.json();
                    results.push(data);
                }
            } catch (err) {
                console.error(`Request failed for: ${sentence}`, err);
            }
        }

        displayResults(results);
    });
}

function displayResults(results) {
    const container = document.getElementById("results");
    container.innerHTML = "";
    let improvedResults = [];

    // simplified better than original
    if (improvementMethod === "better") {
        improvedResults = results.filter(item => 
        item.simplified.lint_score >= 36.18 && item.simplified.lint_score <= 50.07
        );
    }

    // simplified between 36.18 and 50.07 (B1)
    if (improvementMethod === "b1") {
        improvedResults = results.filter(item => 
        item.simplified.lint_score >= 36.18 && item.simplified.lint_score <= 50.07
        );
    }

    if (improvedResults.length === 0) {
        container.innerHTML = "De brief is op B1-niveau!<br><br>U dient de brief zelf kritisch na te lezen.";
        return;
    }

    // Initialize Undo Button if it doesn't exist
    setupUndoButton(container);

    // Create cards for each result
    improvedResults.forEach(item => {
        createSuggestionCard(item, container);
    });

    updateUndoButtonState();
}

function setupUndoButton(container) {
    if (document.getElementById("undoBtn")) return;

    const undoBtn = document.createElement("button");
    undoBtn.id = "undoBtn";
    undoBtn.textContent = "Undo last change";
    undoBtn.style.display = "block";
    undoBtn.style.marginBottom = "15px";

    undoBtn.addEventListener("click", async () => {
        if (undoStack.length === 0) return;

        const lastChange = undoStack.pop();

        if (lastChange.type === "replace") {
            // 1. Revert text in Word
            await replaceInWord(lastChange.fullItem.original.sentence, lastChange.appliedText);
            // 2. Put the card back in the UI
            createSuggestionCard(lastChange.fullItem, container, true);
        } 
        else if (lastChange.type === "deny") {
            // Just put the card back in the UI (no Word changes needed)
            createSuggestionCard(lastChange.fullItem, container, true);
        }

        updateUndoButtonState();
    });

    container.parentElement.insertBefore(undoBtn, container);
}

function createSuggestionCard(item, container, prepend = false) {
    const div = document.createElement("div");
    div.className = "result-card";
    div.style.border = "1px solid #ccc";
    div.style.padding = "10px";
    div.style.marginBottom = "10px";

    div.innerHTML = `
        <div><strong>Original:</strong> ${item.original.sentence}</div>
        <div>Lint Score: ${item.original.lint_score}</div>
        <hr>
        <div><strong>Simplified:</strong> <span class="simplified-text">${item.simplified.sentence}</span></div>
        <div>Lint Score: ${item.simplified.lint_score}</div>
        <div class="button-group" style="margin-top: 10px;">
            <button class="accept-btn">Accepteren</button>
            <button class="modify-btn">Aanpassen</button>
            <button class="deny-btn">Weigeren</button>
        </div>
    `;

    const acceptBtn = div.querySelector(".accept-btn");
    const modifyBtn = div.querySelector(".modify-btn");
    const denyBtn = div.querySelector(".deny-btn");

    // ACCEPT LOGIC
    acceptBtn.addEventListener("click", async () => {
        const textToApply = item.simplified.sentence;
        await replaceInWord(textToApply, item.original.sentence);

        undoStack.push({ fullItem: item, appliedText: textToApply, type: "replace" });
        updateUndoButtonState();
        div.remove();
    });


    // DENY LOGIC
    // Inside createSuggestionCard...
    denyBtn.addEventListener("click", () => {
        // Push to stack as a 'deny' type
        undoStack.push({ 
            fullItem: item, 
            type: "deny" 
        });
        
        updateUndoButtonState();
        div.remove();
    });

    // MODIFY LOGIC
    modifyBtn.addEventListener("click", () => {
        const btnGroup = div.querySelector(".button-group");
        btnGroup.style.display = "none";

        const modifyDiv = document.createElement("div");
        const input = document.createElement("input");
        input.type = "text";
        input.value = item.simplified.sentence;
        input.style.width = "100%";

        const saveBtn = document.createElement("button");
        saveBtn.textContent = "Opslaan";
        const cancelBtn = document.createElement("button");
        cancelBtn.textContent = "Annuleren";

        modifyDiv.append(input, saveBtn, cancelBtn);
        div.appendChild(modifyDiv);

        saveBtn.addEventListener("click", async () => {
            const customText = input.value;

            await replaceInWord(customText, item.original.sentence);

            undoStack.push({ 
                fullItem: item, 
                appliedText: customText,
                type: "replace" 
            });

            updateUndoButtonState();
            div.remove();
        });

        cancelBtn.addEventListener("click", () => {
            modifyDiv.remove();
            btnGroup.style.display = "block";
        });
    });

    if (prepend) {
        container.prepend(div);
    } else {
        container.appendChild(div);
    }
}

async function replaceInWord(newText, oldText) {
    await Word.run(async (context) => {
        // Word Search API crashes if the string is > 255 characters
        // We trim the search string if it's too long, but this can lead to 'not found'
        // A better way is to use the first 50 and last 50 chars to find the range.
        const searchString = oldText.length > 250 ? oldText.substring(0, 250) : oldText;

        const results = context.document.body.search(searchString);
        results.load("items");
        await context.sync();

        if (results.items.length > 0) {
            // If the search was trimmed, we need to expand the selection to the full original sentence 
            // but for simplicity here, we just replace what we found.
            results.items[0].insertText(newText, Word.InsertLocation.replace);
            await context.sync();
        } else {
            console.warn("Could not find the text in the document. It might have been manually changed.");
        }
    }).catch(error => {
        console.error("Word Error: " + error.code + " - " + error.message);
        if (error.code === "SearchStringInvalidOrTooLong") {
            alert("This sentence is too long for Word to find automatically. Please replace it manually.");
        }
    });
}

function updateUndoButtonState() {
    const undoBtn = document.getElementById("undoBtn");
    if (!undoBtn) return;

    const isEmpty = undoStack.length === 0;
    undoBtn.disabled = isEmpty;
    undoBtn.style.opacity = isEmpty ? 0.5 : 1;
    undoBtn.style.cursor = isEmpty ? "not-allowed" : "pointer";
}