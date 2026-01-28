/* global Office, Word */

// =====================
// CONFIG
// =====================
// Demo settings
let improvementMethod = "b1";
let showLintScores = false;

// Pagination
let currentPage = 1;
const pageSize = 1;
let paginatedResults = [];

// State
let initialized = false;
let exporting = false;
let undoStack = [];

// Mock settings
let useMockSuggestions = true;

// =====================
// OFFICE INIT
// =====================
Office.onReady(() => {
    if (initialized) return;
    initialized = true;

    const oldBtn = document.getElementById("export-btn");
    const newBtn = oldBtn.cloneNode(true);
    oldBtn.replaceWith(newBtn);

    newBtn.addEventListener("click", async () => {
        if (exporting) return;

        exporting = true;
        newBtn.disabled = true;
        newBtn.style.opacity = 0.5;
        newBtn.style.cursor = "not-allowed";

        try {
            await exportSentences();   // waits for ALL sentences
        } finally {
            exporting = false;
            newBtn.disabled = false;
            newBtn.style.opacity = 1;
            newBtn.style.cursor = "pointer";
        }
    });

    setupUndoButton();
});


// =====================
// EXPORT
// =====================
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
                // 🧪 MOCK MODE
                if (useMockSuggestions) {
                    results.push(getMockSuggestion(sentence));
                    continue;
                }

                // 🌐 REAL API
                const response = await fetch("http://127.0.0.1:5000/prompt", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ sentence })
                });

                if (response.ok) {
                    const data = await response.json();

                    data.simplified.sentence = data.simplified.sentence
                        .replace(/\[\[\s*##\s*completed\s*##\s*\]\]/gi, "")
                        .trim();

                    // 🔥 store original AI suggestion (for undo after modify)
                    data.simplified.originalSentence = data.simplified.sentence;

                    results.push(data);
                }

            } catch (err) {
                console.error(err);
            }
        }


        displayResults(results);
    });
}

// =====================
// DISPLAY + PAGINATION
// =====================
function displayResults(results) {
    const container = document.getElementById("results");
    container.innerHTML = "";
    currentPage = 1;

    paginatedResults = results.filter(item =>
        item.simplified.lint_score >= 36.18 &&
        item.simplified.lint_score <= 50.07
    );

    if (paginatedResults.length === 0) {
        container.innerHTML =
            "De brief is op B1-niveau!<br><br>U dient de brief zelf kritisch na te lezen.";
        return;
    }

    renderPaginationControls(container);
    renderCurrentPage();
}

function renderPaginationControls(container) {
    const div = document.createElement("div");
    div.id = "pagination";

    div.innerHTML = `
        <button id="prev-page">← Vorige</button>
        <span id="page-info"></span>
        <button id="next-page">Volgende →</button>
    `;

    container.appendChild(div);

    document.getElementById("prev-page").onclick = () => {
        if (currentPage > 1) {
            currentPage--;
            renderCurrentPage();
        }
    };

    document.getElementById("next-page").onclick = () => {
        const max = Math.ceil(paginatedResults.length / pageSize);
        if (currentPage < max) {
            currentPage++;
            renderCurrentPage();
        }
    };
}

function renderCurrentPage() {
    const container = document.getElementById("results");

    container.querySelectorAll(".result-card").forEach(e => e.remove());

    const start = (currentPage - 1) * pageSize;
    const pageItems = paginatedResults.slice(start, start + pageSize);

    pageItems.forEach(item => createSuggestionCard(item, container));

    const maxPage = Math.ceil(paginatedResults.length / pageSize);
    document.getElementById("page-info").textContent =
        `${currentPage} van ${maxPage}`;

    document.getElementById("prev-page").disabled = currentPage === 1;
    document.getElementById("next-page").disabled = currentPage === maxPage;

    // 🔥 NEW: automatically highlight the current suggestion in Word
    if (pageItems.length === 1) {
        highlightInWord(pageItems[0].original.sentence);
    } else {
        clearSelectionInWord();
    }
}


// =====================
// CARD
// =====================
function createSuggestionCard(item, container) {
    const div = document.createElement("div");
    div.className = "result-card";

    div.innerHTML = `
        <div class="sentence-block original-block">
            <strong class="no-break">Originele zin</strong>
            <div>${item.original.sentence}</div>
        </div>

        <div class="sentence-block suggestion-block">
            <strong class="no-break">AI-suggestie</strong>
            <div class="suggestion-text" contenteditable="false">
                ${item.simplified.sentence}
            </div>

            <div class="action-row" style="margin-top:12px;">
                <button class="accept-btn">Accepteren</button>
                <button class="modify-btn">Aanpassen</button>
                <button class="deny-btn">Weigeren</button>
            </div>

            <div class="edit-row" style="margin-top:12px; display:none;">
                <button class="save-btn">Opslaan</button>
                <button class="cancel-btn">Annuleren</button>
            </div>
        </div>
    `;

    const suggestionEl = div.querySelector(".suggestion-text");
    const actionRow = div.querySelector(".action-row");
    const editRow = div.querySelector(".edit-row");

    const originalSuggestion = item.simplified.sentence;

    // ✅ ACCEPT
    div.querySelector(".accept-btn").onclick = async () => {
        await applySuggestion(item, suggestionEl.innerText);
    };

    // ✏️ MODIFY
    div.querySelector(".modify-btn").onclick = () => {
        suggestionEl.contentEditable = "true";
        suggestionEl.focus();

        actionRow.style.display = "none";
        editRow.style.display = "block";
    };

    // 💾 SAVE
    div.querySelector(".save-btn").onclick = async () => {
        const newText = suggestionEl.innerText.trim();

        item.simplified.sentence = newText;
        await applySuggestion(item, newText);
    };

    // ❌ CANCEL
    div.querySelector(".cancel-btn").onclick = () => {
        suggestionEl.innerText = originalSuggestion;
        suggestionEl.contentEditable = "false";

        editRow.style.display = "none";
        actionRow.style.display = "block";
    };

    // 🚫 DENY
    div.querySelector(".deny-btn").onclick = () => {
        undoStack.push({
            type: "deny",
            item,
            pageIndex: paginatedResults.indexOf(item)
        });

        removeItemFromPagination(item);
    };

    container.appendChild(div);
}

async function applySuggestion(item, textToApply) {
    const originalSuggestion = item.simplified.originalSentence ?? item.simplified.sentence;

    await replaceInWord(textToApply, item.original.sentence);

    undoStack.push({
        type: "replace",
        item,
        previousText: item.original.sentence,
        appliedText: textToApply,
        originalSuggestion,           // 🔥 track AI version
        pageIndex: paginatedResults.indexOf(item)
    });

    removeItemFromPagination(item);
}




// =====================
// UNDO
// =====================
function setupUndoButton() {
    const btn = document.getElementById("undo-btn");

    btn.onclick = async () => {
        if (undoStack.length === 0) return;

        const last = undoStack.pop();

        // Restore Word text
        if (last.type === "replace") {
            await replaceInWord(last.previousText, last.appliedText);
            await highlightInWord(last.previousText);

            // 🔥 restore ORIGINAL AI suggestion
            last.item.simplified.sentence = last.originalSuggestion;
        }

        // Restore suggestion into pagination
        if (typeof last.pageIndex === "number") {
            paginatedResults.splice(last.pageIndex, 0, last.item);
            currentPage = last.pageIndex + 1;
        }

        renderCurrentPage();
        updateUndoButtonState();
    };


    updateUndoButtonState();
}


function updateUndoButtonState() {
    const btn = document.getElementById("undo-btn");
    btn.disabled = undoStack.length === 0;
}

// =====================
// WORD HELPERS
// =====================
async function replaceInWord(newText, oldText) {
    await Word.run(async (context) => {
        const search = oldText.substring(0, 250);
        const results = context.document.body.search(search);
        results.load("items");
        await context.sync();

        if (results.items.length > 0) {
            results.items[0].insertText(newText, Word.InsertLocation.replace);
        }
        await context.sync();
    });
}

async function highlightInWord(text) {
    await Word.run(async (context) => {
        const search = text.substring(0, 250);
        const results = context.document.body.search(search);
        results.load("items");
        await context.sync();

        if (results.items.length > 0) {
            results.items[0].select();
        }
        await context.sync();
    });
}

async function clearSelectionInWord() {
    await Word.run(async (context) => {
        context.document.getSelection().select(Word.SelectionMode.end);
        await context.sync();
    });
}

// --

function removeItemFromPagination(item) {
    const index = paginatedResults.indexOf(item);
    if (index !== -1) {
        paginatedResults.splice(index, 1);
    }

    const container = document.getElementById("results");

    // 🔥 NEW: no suggestions left → show B1 message
    if (paginatedResults.length === 0) {
        container.innerHTML =
            "De brief is op B1-niveau!<br><br>U dient de brief zelf kritisch na te lezen.";

        clearSelectionInWord();
        updateUndoButtonState();
        return;
    }

    const maxPage = Math.ceil(paginatedResults.length / pageSize);
    currentPage = Math.min(currentPage, maxPage);

    renderCurrentPage();
    updateUndoButtonState();
}


function getMockSuggestion(sentence) {
    const simplified = sentence
        .replace(/,/g, "")
        .replace(/\b(daarom|echter|desondanks)\b/gi, "omdat")
        .replace(/\s+/g, " ")
        .trim() + ".";

    return {
        original: { sentence },
        simplified: {
            sentence: simplified,
            originalSentence: simplified, // 🔥 keep pristine copy
            lint_score: 42
        }
    };
}

