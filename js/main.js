document.addEventListener('DOMContentLoaded', () => {

  // ===== 1. Footer Year =====
  const yearEl = document.getElementById("year");
  if (yearEl) yearEl.textContent = new Date().getFullYear();

// ===== 2. Support Dropdown Toggle =====
const supportBtn = document.querySelector('.support-btn');
const supportDropdown = document.getElementById('supportDropdown');

if (supportBtn && supportDropdown) {
  // Toggle dropdown on button click
  function toggleSupport() {
    supportDropdown.style.display = supportDropdown.style.display === 'block' ? 'none' : 'block';
  }

  supportBtn.addEventListener('click', (e) => {
    e.stopPropagation(); // Prevent immediate close
    toggleSupport();
  });

  // Close dropdown if clicked outside
  window.addEventListener('click', (e) => {
    if (!e.target.closest('.support-btn') && !e.target.closest('#supportDropdown')) {
      supportDropdown.style.display = 'none';
    }
  });
}

  // ===== 3. File Size Fetch =====
  const fileSizeSpan = document.getElementById('file-size');
  const fileUrl = "https://github.com/Ahmedjawad123/Zamini_Converter/releases/download/v1.0.0/Executable_file_.Zamini_Converter_v1.0.0.rar";

  if (fileSizeSpan) {
    fetch(fileUrl, { method: 'HEAD' })
      .then(resp => {
        const size = resp.headers.get('content-length');
        if (size) fileSizeSpan.textContent = (size / (1024 * 1024)).toFixed(2) + " MB";
        else fileSizeSpan.textContent = "N/A";
      })
      .catch(err => {
        console.error('File size fetch error:', err);
        fileSizeSpan.textContent = "N/A";
      });
  }

  // ===== 4. Initialize EmailJS =====
  if (typeof emailjs !== "undefined") {
    emailjs.init('DhW4bXmuP0VP2d8bF'); // Your Public Key
  } else {
    console.error("EmailJS not loaded!");
  }

  // ===== 5. Feedback Form =====
  const feedbackForm = document.getElementById('contactForm');
  if (feedbackForm) {
    let statusEl = document.getElementById('feedback-status');
    if (!statusEl) {
      statusEl = document.createElement('div');
      statusEl.id = 'feedback-status';
      statusEl.style.marginTop = "8px";
      statusEl.style.color = "green";
      feedbackForm.appendChild(statusEl);
    }

    feedbackForm.addEventListener('submit', (e) => {
      e.preventDefault();
      statusEl.textContent = "Sending...";
      console.log("Sending feedback...");

      const templateParams = {
        software: feedbackForm.software.value || "Not selected",
        name: feedbackForm.name.value || "Anonymous",
        email: feedbackForm.email.value || "Not provided",
        message: feedbackForm.message.value
      };

      emailjs.send('zamini_musafir', 'template_yz15x2d', templateParams)
        .then(response => {
          console.log("EmailJS success:", response);
          statusEl.textContent = "Feedback sent successfully! Thank you.";
          feedbackForm.reset();
        })
        .catch(err => {
          console.error("EmailJS error:", err);
          statusEl.textContent = "Oops! Something went wrong. Check console.";
        });
    });
  }

});
