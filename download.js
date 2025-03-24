// download.js

document.addEventListener("DOMContentLoaded", async function () {
  const typewriterElement = document.querySelector(".typewriter");
  if (typewriterElement) {
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
  console.log("Download page loaded");

  // Get the URL parameters
  const urlParams = new URLSearchParams(window.location.search);
  const filesData = [];

  // Loop through all URL parameters and extract file data
  urlParams.forEach((value, key) => {
    if (key === "name") {
      const name = decodeURIComponent(value);
      const type = decodeURIComponent(urlParams.get("type"));
      const sessionKey = decodeURIComponent(urlParams.get("sessionKey"));

      filesData.push({ name, type, sessionKey });
    }
  });

  if (filesData.length > 0) {
    for (const fileData of filesData) {
      try {
        await processFileData(fileData);
      } catch (error) {
        console.error("Error processing file:", fileData.name, error);
        alert(`❌ Error processing file: ${fileData.name}: ${error.message}`);
      }
    }
  } else {
    console.warn("No files were uploaded or data was lost.");
    alert("No files were uploaded or data was lost. Please try again.");
    window.location.href = "index.html"; // Redirect back if no files
  }

  async function processFileData(fileData) {
    const { name, type, sessionKey } = fileData;

    // Retrieve the ArrayBuffer from session storage
    const base64String = sessionStorage.getItem(sessionKey);
    if (!base64String) {
      throw new Error(`Data missing from session storage for ${name}`);
    }
    const data = base64ToArrayBuffer(base64String); // Convert base64 to ArrayBuffer

    sessionStorage.removeItem(sessionKey); // Clear session storage immediately

    console.log("Processing file:", name, type);
    try {
      if (type.includes("text/plain") || name.endsWith(".txt")) {
        await processTextFileData(data);
      } else if (
        type.includes(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ) ||
        name.endsWith(".docx")
      ) {
        await processDocxFileData(data);
      } else if (type.includes("application/pdf") || name.endsWith(".pdf")) {
        await processPdfFileData(data);
      } else {
        throw new Error(
          "❌ Invalid file type. Please upload a .txt, .docx, or .pdf file."
        );
      }
    } catch (error) {
      console.error(`Error processing ${name}:`, error);
      throw error; // Re-throw to be caught in the outer loop
    }
  }

  async function sendToAPI(extractedText) {
    console.log("📤 Sending extracted text to API...");
    try {
      let response = await fetch(
        "https://flask-summarizer-1-khc9.onrender.com/generate-ppt",
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: extractedText }),
        }
      );

      if (!response.ok)
        throw new Error(
          `Server Error: ${response.status} - ${response.statusText}`
        );

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
      alert("❌ API Request Failed:" + error);
      throw error; // Re-throw to be handled in processFileData
    }
  }

  async function processTextFileData(data) {
    const decoder = new TextDecoder();
    const text = decoder.decode(data);
    await sendToAPI(text);
  }

  async function processDocxFileData(data) {
    try {
      const result = await mammoth.extractRawText({ arrayBuffer: data });
      await sendToAPI(result.value);
    } catch (error) {
      console.error("Error in Mammoth.js:", error);
      throw error;
    }
  }

  async function processPdfFileData(data) {
    try {
      pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js";
      const pdf = await pdfjsLib.getDocument({ data: data }).promise;
      let text = "";

      for (let i = 1; i <= pdf.numPages; i++) {
        let page = await pdf.getPage(i);
        let textContent = await page.getTextContent();
        let pageText = textContent.items
          .map((item) => item.str.trim())
          .join(" ");
        text += pageText + "\n\n";
      }

      await sendToAPI(text.trim());
    } catch (error) {
      console.error("Error processing PDF:", error);
      throw error;
    }
  }

  // Function to convert base64 string to ArrayBuffer
  function base64ToArrayBuffer(base64) {
    const binary_string = atob(base64);
    const len = binary_string.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes.buffer;
  }

  // Function to log API errors with more detail
  function logAPIError(error) {
    console.error("❌ API Request Failed:", error);
    if (error instanceof Error) {
      console.error("Stack trace:", error.stack);
    }
    if (error.response) {
      console.error("Response status:", error.response.status);
      console.error("Response headers:", error.response.headers);
      console.error("Response data:", error.response.data);
    }
  }

  async function sendToAPI(extractedText) {
    console.log("📤 Sending extracted text to API...");
    try {
      let response = await fetch(
        "https://flask-summarizer-1-khc9.onrender.com/generate-ppt",
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: extractedText }),
        }
      );

      if (!response.ok) {
        // Attempt to read the error message from the response body
        let errorMessage = `Server Error: ${response.status} - ${response.statusText}`;
        try {
          const errorData = await response.json(); // Try to parse as JSON
          errorMessage = errorData.message || errorMessage; // Use a specific error message if available
        } catch (parseError) {
          console.warn("Could not parse JSON error from response:", parseError);
        }
        throw new Error(errorMessage);
      }

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
      logAPIError(error);
      alert("❌ API Request Failed:" + error.message);
      throw error; // Re-throw to be handled in processFileData
    }
  }
});
