console.log("📢 PDF.js version:", pdfjsLib.version);

// ✅ Load required libraries dynamically
async function loadLibrary(url, globalVar) {
    return new Promise((resolve, reject) => {
        if (window[globalVar]) return resolve();

        let script = document.createElement("script");
        script.src = url;
        script.async = true;

        script.onload = () => resolve();
        script.onerror = () => reject(new Error(`Failed to load ${globalVar}`));

        document.head.appendChild(script);
    });
}

document.addEventListener("DOMContentLoaded", async () => {
    console.log("🚀 Document Ready!");
    await loadLibrary("https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.9.0/mammoth.browser.min.js", "mammoth");
    await loadLibrary("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js", "pdfjsLib");

    document.querySelector(".upbutton")?.addEventListener("click", handleFileUpload);
});

// ✅ Handle file upload
function handleFileUpload() {
    let fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = ".txt,.docx,.pdf";
    fileInput.style.display = "none";

    document.body.appendChild(fileInput);
    fileInput.click();

    fileInput.addEventListener("change", async () => {
        if (fileInput.files.length) {
            let file = fileInput.files[0];
            let fileType = file.name.split('.').pop().toLowerCase();

            if (fileType === "txt") processTextFile(file);
            else if (fileType === "docx") await processDocxFile(file);
            else if (fileType === "pdf") await processPdfFile(file);
            else alert("❌ Invalid file type. Please upload a .txt, .docx, or .pdf file.");
        }
        fileInput.remove();
    });
}

// ✅ Send extracted text to Flask API for PPT generation
async function sendToAPI(extractedText) {
    console.log("📤 Sending extracted text to API...");
    try {
        let response = await fetch("https://flask-summarizer-1-khc9.onrender.com/generate-ppt", {  // ✅ Replace with actual Render API URL
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ text: extractedText })
        });

        if (!response.ok) throw new Error(`Server Error: ${response.status}`);

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);

        // ✅ Auto-download the PPT file
        const a = document.createElement("a");
        a.href = url;
        a.download = "Generated_Summary_Presentation.pptx";
        document.body.appendChild(a);
        a.click();
        a.remove();

        console.log("✅ PPT file downloaded successfully!");
    } catch (error) {
        console.error("❌ API Request Failed:", error);
    }
}

// ✅ Process TXT file
function processTextFile(file) {
    const reader = new FileReader();
    reader.readAsText(file);
    reader.onload = () => sendToAPI(reader.result);
    reader.onerror = () => console.error("❌ Error reading TXT file.");
}

// ✅ Process DOCX file
async function processDocxFile(file) {
    let reader = new FileReader();
    reader.readAsArrayBuffer(file);

    reader.onload = function (event) {
        let arrayBuffer = event.target.result;
        mammoth.extractRawText({ arrayBuffer })
            .then(result => sendToAPI(result.value))
            .catch(error => console.error("❌ Error in Mammoth.js:", error));
    };
}

// ✅ Process PDF file
async function processPdfFile(file) {
    let reader = new FileReader();
    reader.readAsArrayBuffer(file);

    reader.onload = async function () {
        pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js";
        const pdfData = new Uint8Array(reader.result);
        const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;
        let text = '';

        for (let i = 1; i <= pdf.numPages; i++) {
            let page = await pdf.getPage(i);
            let textContent = await page.getTextContent();
            let pageText = textContent.items.map(item => item.str.trim()).join(" ");
            text += pageText + "\n\n";
        }

        sendToAPI(text.trim());
    };
}