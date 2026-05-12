Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("ask-btn").onclick = runGemini;
        
        // FITUR AUTO-SAVE API KEY: Load saat aplikasi dibuka
        const savedKey = localStorage.getItem("gemini_api_key_v1");
        if (savedKey) {
            document.getElementById("api-key").value = savedKey;
        }
        
        // Simpan setiap kali user mengetik/paste API Key baru
        document.getElementById("api-key").addEventListener('input', function() {
            localStorage.setItem("gemini_api_key_v1", this.value);
        });
    }
});

async function runGemini() {
    const prompt = document.getElementById("prompt").value;
    const apiKey = document.getElementById("api-key").value;
    const responseDiv = document.getElementById("response");

    // Validasi input
    if (!apiKey) {
        responseDiv.innerHTML = "<span style='color: #ff7b72;'>[ERROR] API Key tidak ditemukan. Harap isi di bawah!</span>";
        return;
    }
    if (!prompt) {
        responseDiv.innerHTML = "<span style='color: #d2a8ff;'>[WARN] Instruksi masih kosong.</span>";
        return;
    }

    responseDiv.innerHTML = "<span style='color: var(--accent-color);'>Memproses permintaan data... ⏳</span>";

    try {
        // Ambil konteks sederhana dari file Excel (Nama Sheet)
        let excelContext = "Struktur Workbook:\n";
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            for (let sheet of sheets.items) {
                excelContext += `- Sheet: ${sheet.name}\n`;
            }
        });

        // Panggil fungsi API dan tangkap responnya
        const result = await callGeminiAPI(apiKey, prompt, excelContext);
        responseDiv.innerText = result;

    } catch (error) {
        responseDiv.innerHTML = `<span style='color: #ff7b72;'>[SYS_ERROR] Gagal membaca Excel: ${error.message}</span>`;
    }
}

async function callGeminiAPI(key, userPrompt, context) {
    // FIX: Menggunakan model 'gemini-1.5-pro-latest' agar dikenali oleh server v1beta
    const url =`https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key=${key}`;
    
    const body = {
        contents: [{
            parts: [{ text: `${context}\n\nInstruksi User: ${userPrompt}` }]
        }]
    };

    try {
        const response = await fetch(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json" 
            },
            body: JSON.stringify(body)
        });

        const data = await response.json();

        // CEK APAKAH GOOGLE MENGEMBALIKAN ERROR
        if (!response.ok) {
            if (data.error && data.error.message) {
                return `[API_REJECTED] ${data.error.message}`;
            }
            return `[HTTP_ERROR] Kode Status: ${response.status}`;
        }

        // BACA JAWABAN JIKA BERHASIL
        if (data.candidates && data.candidates.length > 0) {
            return data.candidates[0].content.parts[0].text;
        } else {
            return "[WARN] Respons kosong dari AI.";
        }
    } catch (err) {
        return `[NETWORK_ERROR] Gagal menyambung ke server Google: ${err.message}`;
    }
}