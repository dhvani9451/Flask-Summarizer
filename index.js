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

// ✅ Handle file upload (Supports multiple files)
function handleFileUpload() {
    let fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = ".txt,.docx,.pdf";
    fileInput.multiple = true;  // ✅ Allow multiple file uploads
    fileInput.style.display = "none";

    document.body.appendChild(fileInput);
    fileInput.click();

    fileInput.addEventListener("change", async () => {
        let extractedTexts = {};  // ✅ Store extracted text as { filename: text }

        for (let file of fileInput.files) {
            let fileType = file.name.split('.').pop().toLowerCase();
            let textContent = "";

            if (fileType === "txt") {
                textContent = await processTextFile(file);
            } else if (fileType === "docx") {
                textContent = await processDocxFile(file);
            } else if (fileType === "pdf") {
                textContent = await processPdfFile(file);
            } else {
                alert("❌ Invalid file type. Please upload a .txt, .docx, or .pdf file.");
                continue;
            }

            extractedTexts[file.name] = textContent;  // ✅ Store file name and its content
        }

        if (Object.keys(extractedTexts).length > 0) sendToAPI(extractedTexts);
        fileInput.remove();
    });
}

// ✅ Send extracted text to Flask API for PPT generation
async function sendToAPI(extractedTexts) {
  console.log("📤 Sending extracted texts to API...");

  try {
      let response = await fetch("https://flask-summarizer-1-khc9.onrender.com/generate-ppt", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ texts: extractedTexts }),
          mode: "cors", // ✅ Explicitly enable CORS
      });

      if (!response.ok) throw new Error(`Server Error: ${response.status}`);

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = "Generated_Summary_Presentation.pptx";
      document.body.appendChild(a);
      a.click();
      a.remove();

      console.log("✅ PPT file downloaded successfully!");
  } catch (error) {
      console.error("❌ API Request Failed:", error);
      alert("There was an error generating the PPT. Try uploading fewer files.");
  }
}


// ✅ Process TXT file
async function processTextFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsText(file);
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject("❌ Error reading TXT file.");
    });
}

// ✅ Process DOCX file
async function processDocxFile(file) {
    return new Promise((resolve, reject) => {
        let reader = new FileReader();
        reader.readAsArrayBuffer(file);

        reader.onload = function (event) {
            let arrayBuffer = event.target.result;
            mammoth.extractRawText({ arrayBuffer })
                .then(result => resolve(result.value))
                .catch(error => reject("❌ Error in Mammoth.js: " + error));
        };

        reader.onerror = () => reject("❌ Error reading DOCX file.");
    });
}

// ✅ Process PDF file
async function processPdfFile(file) {
    return new Promise((resolve, reject) => {
        let reader = new FileReader();
        reader.readAsArrayBuffer(file);

        reader.onload = async function () {
            try {
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

                resolve(text.trim());
            } catch (error) {
                reject("❌ Error processing PDF file: " + error);
            }
        };

        reader.onerror = () => reject("❌ Error reading PDF file.");
    });
}
