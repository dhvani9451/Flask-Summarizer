//JS for index page animation

var TxtType = function (el, toRotate, period) {
  this.toRotate = toRotate;
  this.el = el;
  this.loopNum = 0;
  this.period = parseInt(period, 10) || 2000;
  this.txt = "";
  this.tick();
  this.isDeleting = false;
};

TxtType.prototype.tick = function () {
  var i = this.loopNum % this.toRotate.length;
  var fullTxt = this.toRotate[i];

  if (this.isDeleting) {
    this.txt = fullTxt.substring(0, this.txt.length - 1);
  } else {
    this.txt = fullTxt.substring(0, this.txt.length + 1);
  }

  this.el.innerHTML = '<span class="wrap">' + this.txt + "</span>";

  var that = this;
  var delta = 200 - Math.random() * 100;

  if (this.isDeleting) {
    delta /= 2;
  }

  if (!this.isDeleting && this.txt === fullTxt) {
    delta = this.period;
    this.isDeleting = true;
  } else if (this.isDeleting && this.txt === "") {
    this.isDeleting = false;
    this.loopNum++;
    delta = 500;
  }

  setTimeout(function () {
    that.tick();
  }, delta);
};

window.onload = function () {
  var elements = document.getElementsByClassName("typewrite");
  for (var i = 0; i < elements.length; i++) {
    var toRotate = elements[i].getAttribute("data-type");
    var period = elements[i].getAttribute("data-period");
    if (toRotate) {
      new TxtType(elements[i], JSON.parse(toRotate), period);
    }
  }

  // INJECT CSS to change cursor color
  var css = document.createElement("style");
  css.type = "text/css";
  css.innerHTML =
    ".typewrite > .wrap { border-right: 0.08em solid rgb(85, 79, 255);}"; // Change cursor color to orange (#FF5733)
  document.body.appendChild(css);
};

document.addEventListener("DOMContentLoaded", function () {
  // Check if the current page is "download.html"
  if (window.location.pathname.includes("download.html")) {
    const typewriterElement = document.querySelector(".typewriter");

    if (!typewriterElement) {
      console.error(
        "❌ Error: .typewriter element not found in download.html!"
      );
      return;
    }

    const textToType = JSON.parse(
      typewriterElement.getAttribute("data-type")
    )[0];
    let index = 0;

    function type() {
      if (index < textToType.length) {
        typewriterElement.textContent += textToType.charAt(index);
        index++;
        setTimeout(type, 100);
      }
    }

    type();
  }
});

// index.js - Handles file uploads, text extraction, and API request

console.log("📢 PDF.js version:", pdfjsLib?.version || "Not loaded");

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
  await loadLibrary(
    "https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.9.0/mammoth.browser.min.js",
    "mammoth"
  );
  await loadLibrary(
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js",
    "pdfjsLib"
  );
  document
    .querySelector(".upbutton")
    ?.addEventListener("click", handleFileUpload);
});

// ✅ Handle file upload (Multiple files allowed)
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
        await processFile(file); // Ensure files are processed before redirecting
      }
      console.log("✅ All files processed. Redirecting...");
      window.location.href = "download.html";
    }
  });
}

// ✅ Process file based on type
async function processFile(file) {
  let fileType = file.name.split(".").pop().toLowerCase();
  console.log("📂 Processing File:", file.name);

  if (fileType === "txt") await processTextFile(file);
  else if (fileType === "docx") await processDocxFile(file);
  else if (fileType === "pdf") await processPdfFile(file);
  else
    alert("❌ Invalid file type. Please upload a .txt, .docx, or .pdf file.");
}

// ✅ Send extracted text to Flask API for PPT generation
async function sendToAPI(extractedText) {
  if (!extractedText || extractedText.trim() === "") {
    console.error("❌ No extracted text to send!");
    return;
  }
  console.log(
    "📤 Sending extracted text to API:",
    extractedText.substring(0, 100),
    "..."
  );

  try {
    let response = await fetch(
      "https://flask-summarizer-1-khc9.onrender.com/generate-ppt",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text: extractedText }),
      }
    );

    console.log("📨 API Response Status:", response.status);
    if (!response.ok) throw new Error(`Server Error: ${response.status}`);

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);

    sessionStorage.setItem("pptDownloadLink", url);
    console.log("✅ PPT link stored in sessionStorage:", url);

    window.location.href = "download.html";
  } catch (error) {
    console.error("❌ API Request Failed:", error);
  }
}

// ✅ Process TXT file
async function processTextFile(file) {
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
    mammoth
      .extractRawText({ arrayBuffer })
      .then((result) => sendToAPI(result.value))
      .catch((error) => console.error("❌ Error in Mammoth.js:", error));
  };
}

// ✅ Process PDF file
async function processPdfFile(file) {
  let reader = new FileReader();
  reader.readAsArrayBuffer(file);

  reader.onload = async function () {
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
    mammoth
      .extractRawText({ arrayBuffer })
      .then((result) => sendToAPI(result.value))
      .catch((error) => console.error("❌ Error in Mammoth.js:", error));
  };
}

// ✅ Process PDF file
async function processPdfFile(file) {
  let reader = new FileReader();
  reader.readAsArrayBuffer(file);

  reader.onload = async function () {
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
