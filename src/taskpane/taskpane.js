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
        await exportSentences();
        exporting = false;
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

function displayResults(results) {
    const container = document.getElementById("results");
    container.innerHTML = ""; // Clear previous

    results.forEach((item, idx) => {
        const div = document.createElement("div");
        div.style.border = "1px solid #ccc";
        div.style.padding = "10px";
        div.style.marginBottom = "8px";

        div.innerHTML = `
            <strong>Sentence ${idx + 1}:</strong><br>
            <strong>Original:</strong> ${item.original.sentence}<br>
            Lint Score: ${item.original.lint_score} | Valid: ${item.original.valid}<br>
            <strong>Simplified:</strong> ${item.simplified.sentence}<br>
            Lint Score: ${item.simplified.lint_score} | Valid: ${item.simplified.valid}<br>
            <strong>Improvement:</strong> ${item.improvement}
        `;
        container.appendChild(div);
    });
}
