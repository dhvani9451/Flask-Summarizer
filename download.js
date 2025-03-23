// download.js - Handles automatic PPT download in download.html

document.addEventListener("DOMContentLoaded", function () {
  console.log("📂 Running download.js");

  // ✅ Typewriter effect
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
  } else {
    console.error("❌ Error: .typewriter element not found in download.html!");
  }

  // ✅ Get the stored PPT link
  const pptLink = sessionStorage.getItem("pptDownloadLink");

  if (pptLink) {
    console.log("📥 PPT file found in sessionStorage:", pptLink);
    autoDownloadPPT(pptLink);
  } else {
    console.error("❌ No PPT file found in sessionStorage.");
    alert("❌ No PPT file available. Please re-upload your document.");
  }
});

// ✅ Function to automatically download the PPT file
function autoDownloadPPT(url) {
  const a = document.createElement("a");
  a.href = url;
  a.download = "Generated_Summary_Presentation.pptx";
  document.body.appendChild(a);
  a.click();
  a.remove();
}
