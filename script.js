Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("ask-btn").onclick = runGemini;
    }
});

async function runGemini() {
    const prompt = document.getElementById("prompt").value;
    const apiKey = document.getElementById("api-key").value;
    const responseDiv = document.getElementById("response");

    if (!apiKey) {
        alert("Masukkan API Key dulu di bagian bawah!");
        return;
    }

    responseDiv.innerText = "Sedang berpikir...";

    try {
        // Logika Membaca Konteks Excel (Nama Sheet & Header)
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            let excelContext = "Struktur Workbook saya:\n";
            for (let sheet of sheets.items) {
                excelContext += `- Sheet: ${sheet.name}\n`;
            }

            // Panggil API Gemini
            const result = await callGeminiAPI(apiKey, prompt, excelContext);
            responseDiv.innerText = result;
        });
    } catch (error) {
        responseDiv.innerText = "Error: " + error.message;
    }
}

async function callGeminiAPI(key, userPrompt, context) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`;
    
    const body = {
        contents: [{
            parts: [{ text: `${context}\n\nPertanyaan User: ${userPrompt}` }]
        }]
    };

    const response = await fetch(url, {
        method: "POST",
        body: JSON.stringify(body)
    });

    const data = await response.json();
    return data.candidates[0].content.parts[0].text;
}