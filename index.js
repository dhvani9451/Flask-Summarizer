console.log("📢 PDF.js is loading...");

// ✅ Function to dynamically load external libraries
async function loadLibrary(url, globalVar) {
  return new Promise((resolve, reject) => {
    if (window[globalVar]) return resolve();

    let script = document.createElement("script");
    script.src = url;
    script.async = true;

    script.onload = () => {
      console.log(`✅ ${globalVar} loaded successfully`);
      resolve();
    };
    script.onerror = () => reject(new Error(`❌ Failed to load ${globalVar}`));

    document.head.appendChild(script);
  });
}

document.addEventListener("DOMContentLoaded", async () => {
  console.log("🚀 Document Ready!");

  try {
    await loadLibrary(
      "https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.9.0/mammoth.browser.min.js",
      "mammoth"
    );
    await loadLibrary(
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js",
      "pdfjsLib"
    );

    if (!window.pdfjsLib) {
      throw new Error("⚠️ PDF.js failed to load.");
    }

    console.log("📢 PDF.js version:", pdfjsLib.version);

    document
      .querySelector(".upbutton")
      ?.addEventListener("click", handleFileUpload);

    setupDragAndDrop();
  } catch (error) {
    console.error("❌ Error loading libraries:", error);
  }
});

// ✅ Function to process uploaded files
function handleFileUpload() {
  let fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = ".txt,.docx,.pdf";
  fileInput.multiple = true;
  fileInput.style.display = "none";

  document.body.appendChild(fileInput);
  fileInput.click();

  fileInput.addEventListener("change", async () => {
    if (fileInput.files.length) {
      for (let file of fileInput.files) {
        await processFile(file);
      }
    }
  });
}

// ✅ Function to process different file types
async function processFile(file) {
  let fileType = file.name.split(".").pop().toLowerCase();

  if (fileType === "txt") await processTextFile(file);
  else if (fileType === "docx") await processDocxFile(file);
  else if (fileType === "pdf") await processPdfFile(file);
  else alert("❌ Invalid file type. Please upload a .txt, .docx, or .pdf file.");
}

// ✅ Process PDF file (Fixed)
async function processPdfFile(file) {
  let reader = new FileReader();
  reader.readAsArrayBuffer(file);

  reader.onload = async function () {
    if (!window.pdfjsLib) {
      console.error("❌ PDF.js is not loaded.");
      return;
    }

    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js";
    
    const pdfData = new Uint8Array(reader.result);
    const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;
    let text = "";

    for (let i = 1; i <= pdf.numPages; i++) {
      let page = await pdf.getPage(i);
      let textContent = await page.getTextContent();
      let pageText = textContent.items.map((item) => item.str.trim()).join(" ");
      text += pageText + "\n\n";
    }

    sendToAPI(text.trim());
  };
}
