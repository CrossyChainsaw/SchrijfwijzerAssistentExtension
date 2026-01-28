/* global Office, Word */

// =====================
// CONFIG
// =====================
let improvementMethod = "b1";
let showLintScores = true;

let useMockSuggestions = true;

// Pagination
let currentPage = 1;
const pageSize = 1;
let paginatedResults = [];

// State
let initialized = false;
let exporting = false;
let undoStack = [];

// =====================
// OFFICE INIT
// =====================
Office.onReady(() => {
    if (initialized) return;
    initialized = true;

    const oldBtn = document.getElementById("export-btn");
    const newBtn = oldBtn.cloneNode(true);
    oldBtn.replaceWith(newBtn);

    newBtn.addEventListener("click", exportAllSentences);

    setupUndoButton();
});

// =====================
// EXPORT SENTENCES
// =====================
async function exportAllSentences() {
    if (exporting) return;

    exporting = true;
    const btn = document.getElementById("export-btn");

    disableAllButtons([btn]); 
    setButtonState(btn, true);

    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.load("text");
            await context.sync();

            // Split document into sentences
            const sentences = body.text.split(/(?<=[.!?])\s+/);
            const results = [];

            for (const sentence of sentences) {
                const suggestion = useMockSuggestions
                    ? getMockSuggestion(sentence)
                    : await fetchSuggestion(sentence);

                if (suggestion) {
                    results.push({
                        original: suggestion.original,
                        simplified: suggestion.simplified
                    });
                }
            }

            // Display suggestions in the UI
            displayResults(results);
        });
    } catch (err) {
        console.error(err);
    } finally {
        exporting = false;
        setButtonState(btn, false);
        enableAllButtons(); 
    }
}



function setButtonState(button, disabled) {
    button.disabled = disabled;
    button.style.opacity = disabled ? 0.5 : 1;
    button.style.cursor = disabled ? "not-allowed" : "pointer";
}

async function fetchSuggestion(sentence) {
    const response = await fetch("http://127.0.0.1:5000/prompt", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sentence })
    });

    if (!response.ok) return null;

    const data = await response.json();
    data.simplified.sentence = data.simplified.sentence
        .replace(/\[\[\s*##\s*completed\s*##\s*\]\]/gi, "")
        .trim();

    data.simplified.originalSentence = data.simplified.sentence;

    // Add lint_score to original
    data.original = data.original || {};
    data.original.sentence = sentence;
    data.original.lint_score = data.original.lint_score ?? data.simplified.lint_score;

    return data;
}


// =====================
// DISPLAY + PAGINATION
// =====================
function displayResults(results) {
    paginatedResults = results.filter(item =>
        item.simplified.lint_score >= 36.18 &&
        item.simplified.lint_score <= 50.07
    );

    const container = document.getElementById("results");
    container.innerHTML = "";

    if (!paginatedResults.length) {
        container.innerHTML = "De brief is op B1-niveau!<br><br>U dient de brief zelf kritisch na te lezen.";
        return;
    }

    currentPage = 1;
    renderPagination(container);
    renderCurrentPage();
}

function renderPagination(container) {
    // Check if pagination already exists
    let div = document.getElementById("pagination");
    if (!div) {
        div = document.createElement("div");
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
            const maxPage = Math.ceil(paginatedResults.length / pageSize);
            if (currentPage < maxPage) {
                currentPage++;
                renderCurrentPage();
            }
        };
    }
}


function renderCurrentPage() {
    const container = document.getElementById("results");
    container.querySelectorAll(".result-card").forEach(e => e.remove());

    const start = (currentPage - 1) * pageSize;
    const pageItems = paginatedResults.slice(start, start + pageSize);

    pageItems.forEach(item => createSuggestionCard(item, container));

    const maxPage = Math.ceil(paginatedResults.length / pageSize);
    document.getElementById("page-info").textContent = `${currentPage} van ${maxPage}`;
    document.getElementById("prev-page").disabled = currentPage === 1;
    document.getElementById("next-page").disabled = currentPage === maxPage;

    if (pageItems.length === 1) highlightInWord(pageItems[0].original.sentence);
    else clearSelectionInWord();
}

// =====================
// CARD
// =====================
function createSuggestionCard(item, container) {
    const div = document.createElement("div");
    div.className = "result-card";

    // Only show lint score if available
    const lintScoreSimpHtml = item.simplified.lint_score != null
        ? `<div class="lint-score">Lint score: ${item.simplified.lint_score.toFixed(2)}</div>`
        : "";

            // Only show lint score if available
    const lintScoreOrgiHtml = item.original.lint_score != null
        ? `<div class="lint-score">Lint score: ${item.original.lint_score.toFixed(2)}</div>`
        : "";

    div.innerHTML = `
        <div class="sentence-block original-block">
            <strong>Originele zin</strong>
            <div>${item.original.sentence}</div>
            ${lintScoreOrgiHtml}
        </div>

        <div class="sentence-block suggestion-block" style="position: relative;">
            <button class="refresh-btn" style="
                position: absolute;
                top: 4px;
                right: 4px;
                font-size: 0.8em;
                padding: 2px 6px;
                cursor: pointer;
            ">Nieuwe suggestie ↻</button>

            <strong>AI-suggestie</strong>
            <div class="suggestion-text" contenteditable="false">${item.simplified.sentence}</div>
            ${lintScoreSimpHtml}
            <div class="action-row" style="margin-top:12px;">
                <button class="accept-btn">Accepteren</button>
                <button class="modify-btn">Aanpassen</button>
                <button class="deny-btn">Weigeren</button>
            </div>
        </div>
    `;


    container.appendChild(div);

    const buttons = {
        accept: div.querySelector(".accept-btn"),
        modify: div.querySelector(".modify-btn"),
        deny: div.querySelector(".deny-btn"),
        refresh: div.querySelector(".refresh-btn")
    };


    buttons.refresh.onclick = async () => {
        // Disable all buttons except this refresh
        disableAllButtons([buttons.refresh]);

        buttons.refresh.disabled = true;
        buttons.refresh.textContent = "Bezig…";

        try {
            const newSuggestion = useMockSuggestions
                ? getMockSuggestion(item.original.sentence)
                : await fetchSuggestion(item.original.sentence);

            if (newSuggestion) {
                undoStack.push({
                    type: "modify",
                    item,
                    previousSentence: item.simplified.sentence,
                    pageIndex: paginatedResults.indexOf(item)
                });

                item.simplified = newSuggestion.simplified;

                const suggestionDiv = div.querySelector(".suggestion-text");
                suggestionDiv.textContent = item.simplified.sentence;

                const lintDiv = div.querySelector(".sentence-block.suggestion-block .lint-score");
                if (lintDiv) {
                    lintDiv.textContent = `Lint score: ${item.simplified.lint_score.toFixed(2)}`;
                }

                updateUndoButtonState();
            }
        } catch (err) {
            console.error(err);
        } finally {
            buttons.refresh.disabled = false;
            buttons.refresh.textContent = "Nieuwe suggestie ↻";
            enableAllButtons(); // re-enable all buttons
        }
    };




    buttons.accept.onclick = () => applySuggestion(item, item.simplified.sentence);
    buttons.deny.onclick = () => removeItemFromPagination(item);
    buttons.modify.onclick = () => enableModify(item, div);

    buttons.deny.onclick = () => {
    // Push undo info so we can restore the denied item
    undoStack.push({
        type: "deny",
        item,
        pageIndex: paginatedResults.indexOf(item)
    });

    removeItemFromPagination(item);
    updateUndoButtonState();
};

}

// =====================
// MODIFY
// =====================
function enableModify(item, cardDiv) {
    const suggestionDiv = cardDiv.querySelector(".suggestion-text");
    const oldText = item.simplified.sentence;

    const textarea = document.createElement("textarea");
    textarea.value = oldText;
    textarea.style.width = "100%";
    textarea.rows = 4;
    suggestionDiv.replaceWith(textarea);

    const buttonRow = document.createElement("div");
    buttonRow.style.marginTop = "12px";
    buttonRow.style.display = "flex";             // make it flex
    buttonRow.style.justifyContent = "space-between"; // push buttons to opposite ends


    const saveBtn = document.createElement("button");
    saveBtn.className = "save-btn";
    saveBtn.textContent = "Opslaan";

    const cancelBtn = document.createElement("button");
    cancelBtn.className = "cancel-btn";
    cancelBtn.textContent = "Annuleren";

    buttonRow.append(saveBtn, cancelBtn);
    textarea.after(buttonRow);


    toggleCardButtons(cardDiv, false);
    disableAllButtons([]); // disable everything


    saveBtn.onclick = () => {
        const newText = textarea.value.trim();
        if (!newText) return;

        undoStack.push({ type: "modify", item, previousSentence: oldText, pageIndex: paginatedResults.indexOf(item) });
        item.simplified.sentence = newText;
        restoreCardView(cardDiv, item);

        
        enableAllButtons(); // re-enable everything
    };

    cancelBtn.onclick = () => {
        restoreCardView(cardDiv, item);
        enableAllButtons(); // re-enable everything
    };
}

function restoreCardView(cardDiv, item) {
    const newDiv = document.createElement("div");
    newDiv.className = "suggestion-text";
    newDiv.contentEditable = false;
    newDiv.textContent = item.simplified.sentence;

    const textarea = cardDiv.querySelector("textarea");
    const buttonRow = textarea.nextElementSibling;
    textarea.replaceWith(newDiv);
    buttonRow.remove();

    toggleCardButtons(cardDiv, true);
    updateUndoButtonState();
    updatePaginationButtons(); // <-- add this line

}

function toggleCardButtons(cardDiv, enabled) {
    cardDiv.querySelectorAll(".accept-btn, .modify-btn, .deny-btn")
        .forEach(btn => btn.style.display = enabled ? "" : "none");
}

// =====================
// APPLY SUGGESTION
// =====================
async function applySuggestion(item, textToApply) {
    const originalSuggestion = item.simplified.originalSentence ?? item.simplified.sentence;

    await replaceInWord(textToApply, item.original.sentence);

    undoStack.push({
        type: "replace",
        item,
        previousText: item.original.sentence,
        appliedText: textToApply,
        originalSuggestion,
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
        if (!undoStack.length) return;
        const last = undoStack.pop();

        if (last.type === "replace") {
            await replaceInWord(last.previousText, last.appliedText);
            await highlightInWord(last.previousText);
            paginatedResults.splice(last.pageIndex, 0, last.item);
            currentPage = last.pageIndex + 1;
        }

        if (last.type === "modify") {
            last.item.simplified.sentence = last.previousSentence;
            currentPage = last.pageIndex + 1;
        }

        if (last.type === "deny") {
            // Restore the denied item in paginatedResults
            paginatedResults.splice(last.pageIndex, 0, last.item);
            currentPage = last.pageIndex + 1;
        }


        renderCurrentPage();
        updateUndoButtonState();
    };

    updateUndoButtonState();
}

function updateUndoButtonState() {
    document.getElementById("undo-btn").disabled = undoStack.length === 0;
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
        if (results.items.length > 0) results.items[0].insertText(newText, Word.InsertLocation.replace);
        await context.sync();
    });
}

async function highlightInWord(text) {
    await Word.run(async (context) => {
        const search = text.substring(0, 250);
        const results = context.document.body.search(search);
        results.load("items");
        await context.sync();
        if (results.items.length > 0) results.items[0].select();
        await context.sync();
    });
}

async function clearSelectionInWord() {
    await Word.run(async context => {
        context.document.getSelection().select(Word.SelectionMode.end);
        await context.sync();
    });
}

// =====================
// PAGINATION HELPERS
// =====================
function removeItemFromPagination(item) {
    const index = paginatedResults.indexOf(item);
    if (index !== -1) paginatedResults.splice(index, 1);

    const container = document.getElementById("results");
    if (!paginatedResults.length) {
        container.innerHTML = "De brief is op B1-niveau!<br><br>U dient de brief zelf kritisch na te lezen.";
        clearSelectionInWord();
        updateUndoButtonState();
        return;
    }

    currentPage = Math.min(currentPage, Math.ceil(paginatedResults.length / pageSize));
    renderCurrentPage();
    updateUndoButtonState();
}

// =====================
// MOCK
// =====================
function getMockSuggestion(sentence) {
    const simplified = sentence
        .replace(/,/g, "")
        .replace(/\b(daarom|echter|desondanks)\b/gi, "omdat")
        .replace(/\s+/g, " ")
        .trim() + ".";

    return {
        original: {
            sentence,
            lint_score: 50 // or some realistic original lint score
        },
        simplified: {
            sentence: simplified,
            originalSentence: simplified,
            lint_score: 42
        }
    };
}


function disableAllButtons(except = []) {
    const btnSelectors = [
        "#export-btn",
        "#undo-btn",
        "#prev-page",
        "#next-page",
        ".accept-btn",
        ".modify-btn",
        ".deny-btn",
        ".refresh-btn"
    ];

    btnSelectors.forEach(selector => {
        const buttons = document.querySelectorAll(selector);
        buttons.forEach(btn => {
            if (!except.includes(btn)) {
                btn.disabled = true;
                btn.style.opacity = 0.5;
                btn.style.cursor = "not-allowed";
            }
        });
    });
}

function enableAllButtons() {
    const btnSelectors = [
        "#export-btn",
        "#undo-btn",
        "#prev-page",
        "#next-page",
        ".accept-btn",
        ".modify-btn",
        ".deny-btn",
        ".refresh-btn"
    ];

    btnSelectors.forEach(selector => {
        const buttons = document.querySelectorAll(selector);
        buttons.forEach(btn => {
            btn.disabled = false;
            btn.style.opacity = 1;
            btn.style.cursor = "pointer";
        });
    });
}

function updatePaginationButtons() {
    const maxPage = Math.ceil(paginatedResults.length / pageSize);
    document.getElementById("prev-page").disabled = currentPage === 1;
    document.getElementById("next-page").disabled = currentPage === maxPage;
}
