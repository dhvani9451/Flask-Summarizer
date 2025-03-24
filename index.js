<<<<<<< HEAD
// ✅ Typewriter effect JS code
var TxtType = function(el, toRotate, period) {
    this.toRotate = toRotate;
    this.el = el;
    this.loopNum = 0;
    this.period = parseInt(period, 10) || 2000;
    this.txt = '';
    this.tick();
    this.isDeleting = false;
};

TxtType.prototype.tick = function() {
    var i = this.loopNum % this.toRotate.length;
    var fullTxt = this.toRotate[i];

    if (this.isDeleting) {
        this.txt = fullTxt.substring(0, this.txt.length - 1);
    } else {
        this.txt = fullTxt.substring(0, this.txt.length + 1);
    }

    this.el.innerHTML = '<span class="wrap">'+this.txt+'</span>';

    var that = this;
    var delta = 200 - Math.random() * 100;

    if (this.isDeleting) { delta /= 2; }

    if (!this.isDeleting && this.txt === fullTxt) {
        delta = this.period;
        this.isDeleting = true;
    } else if (this.isDeleting && this.txt === '') {
        this.isDeleting = false;
        this.loopNum++;
        delta = 500;
    }

    setTimeout(function() {
        that.tick();
    }, delta);
};

window.onload = function() {
    var elements = document.getElementsByClassName('typewrite');
    for (var i = 0; i < elements.length; i++) {
        var toRotate = elements[i].getAttribute('data-type');
        var period = elements[i].getAttribute('data-period');
        if (toRotate) {
            new TxtType(elements[i], JSON.parse(toRotate), period);
        }
    }

    var css = document.createElement("style");
    css.type = "text/css";
    css.innerHTML = ".typewrite > .wrap { border-right: 0.08em solid rgb(85, 79, 255);}";
    document.body.appendChild(css);
};

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
    setupDragAndDrop(); // ✅ Initialize drag and drop
});

// ✅ Handle file upload (Supports multiple files)
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
            processFiles(fileInput.files);
        }
        fileInput.remove();
    });
}

// ✅ Drag and Drop Support
function setupDragAndDrop() {
    const overlay = document.createElement("div");
    overlay.className = "drag-overlay";
    document.body.appendChild(overlay);

    let dragCounter = 0;

    document.addEventListener("dragenter", (event) => {
        event.preventDefault();
        dragCounter++;
        overlay.classList.add("visible");
    });

    document.addEventListener("dragover", (event) => event.preventDefault());

    document.addEventListener("dragleave", (event) => {
        dragCounter--;
        if (dragCounter === 0) {
            overlay.classList.remove("visible");
        }
    });

    document.addEventListener("drop", (event) => {
        event.preventDefault();
        overlay.classList.remove("visible");
        dragCounter = 0;

        if (event.dataTransfer.files.length) {
            processFiles(event.dataTransfer.files);
        }
    });

    window.addEventListener("dragover", (event) => event.preventDefault());
    window.addEventListener("drop", (event) => event.preventDefault());
}

// ✅ Process multiple files
async function processFiles(files) {
    let extractedTexts = {};

    for (let file of files) {
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

        extractedTexts[file.name] = textContent;
    }

    if (Object.keys(extractedTexts).length > 0) sendToAPI(extractedTexts);
}

// ✅ Send extracted text to Flask API for PPT generation
async function sendToAPI(extractedTexts) {
    console.log("📤 Sending extracted texts to API...");
    try {
        let response = await fetch("https://flask-summarizer-1-khc9.onrender.com/generate-ppt", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ file_texts: extractedTexts }) // ✅ Correct structure
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
    });
}