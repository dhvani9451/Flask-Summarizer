// index.js

//JS for index page animation

var TxtType = function (el, toRotate, period) {
    this.toRotate = toRotate;
    this.el = el;
    this.loopNum = 0;
    this.period = parseInt(period, 10) || 2000;
    this.txt = '';
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

    this.el.innerHTML = '<span class="wrap">' + this.txt + '</span>';

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

    setTimeout(function () {
        that.tick();
    }, delta);
};

window.onload = function () {
    var elements = document.getElementsByClassName('typewrite');
    for (var i = 0; i < elements.length; i++) {
        var toRotate = elements[i].getAttribute('data-type');
        var period = elements[i].getAttribute('data-period');
        if (toRotate) {
            new TxtType(elements[i], JSON.parse(toRotate), period);
        }
    }

    // INJECT CSS to change cursor color
    var css = document.createElement("style");
    css.type = "text/css";
    css.innerHTML = ".typewrite > .wrap { border-right: 0.08em solid rgb(85, 79, 255);}"; // Change cursor color to orange (#FF5733)
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

    setupDragAndDrop(); // ✅ Initialize drag and drop
});

// ✅ Handle file upload (Multiple files allowed)
async function handleFileUpload() {
    let fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = ".txt,.docx,.pdf";
    fileInput.multiple = true; // ✅ Allow multiple file uploads
    fileInput.style.display = "none";

    document.body.appendChild(fileInput);
    fileInput.click();

    fileInput.addEventListener("change", async () => {
        if (fileInput.files.length) {
            const files = Array.from(fileInput.files);
            const fileParams = []; // Array to store file metadata for URL

            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                const fileData = await readFileData(file);

                // Store file data in session storage, using a unique key
                const sessionStorageKey = `fileData_${i}`;
                sessionStorage.setItem(sessionStorageKey, arrayBufferToBase64(fileData));  // Store as base64 string

                // Create metadata for URL
                fileParams.push(`name=${encodeURIComponent(file.name)}&type=${encodeURIComponent(file.type)}&sessionKey=${encodeURIComponent(sessionStorageKey)}`);
            }

            // Construct the URL with file parameters
            const urlParams = fileParams.join('&');
            location.href = `download.html?${urlParams}`;
        }
        fileInput.remove();
    });
}

// Function to read file as ArrayBuffer
function readFileData(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            resolve(e.target.result); // result is an ArrayBuffer
        };
        reader.onerror = (e) => {
            reject(e);
        };
        reader.readAsArrayBuffer(file);
    });
}

// Function to convert ArrayBuffer to base64 string (for session storage)
function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

// ✅ Drag and Drop Feature with Fullscreen Overlay Effect (Multiple files supported)
function setupDragAndDrop() {
    const overlay = document.createElement("div");
    overlay.className = "drag-overlay";
    document.body.appendChild(overlay);

    let dragCounter = 0; // To track multiple drag events

    document.addEventListener("dragenter", (event) => {
        event.preventDefault();
        dragCounter++;
        overlay.classList.add("visible");
    });

    document.addEventListener("dragover", (event) => {
        event.preventDefault();
    });

    document.addEventListener("dragleave", (event) => {
        dragCounter--;
        if (dragCounter === 0) {
            overlay.classList.remove("visible");
        }
    });

    document.addEventListener("drop", async (event) => {  // MARKED AS ASYNC
        event.preventDefault();
        overlay.classList.remove("visible");
        dragCounter = 0;

        if (event.dataTransfer.files.length) {
            const files = Array.from(event.dataTransfer.files);
            const fileParams = []; // Array to store file metadata for URL

            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                const fileData = await readFileData(file);

                // Store file data in session storage, using a unique key
                const sessionStorageKey = `fileData_${i}`;
                sessionStorage.setItem(sessionStorageKey, arrayBufferToBase64(fileData));  // Store as base64 string

                // Create metadata for URL
                fileParams.push(`name=${encodeURIComponent(file.name)}&type=${encodeURIComponent(file.type)}&sessionKey=${encodeURIComponent(sessionStorageKey)}`);
            }

            // Construct the URL with file parameters
            const urlParams = fileParams.join('&');
            location.href = `download.html?${urlParams}`;
        }
    });

    // Prevent default behavior for entire window
    window.addEventListener("dragover", (event) => event.preventDefault());
    window.addEventListener("drop", (event) => event.preventDefault());
}