document.addEventListener('DOMContentLoaded', () => {

  // ===== 1. Footer Year =====
  const yearEl = document.getElementById("year");
  if (yearEl) yearEl.textContent = new Date().getFullYear();

// ===== 2. Support Button & PayPal Popup =====
const supportBtn = document.getElementById('supportBtn');
const paypalPopup = document.getElementById('paypalPopup');
let isPayPalRendered = false;
let selectedAmount = "5"; // default

const quickBtn = document.getElementById('donate-5');
const customInput = document.getElementById('donate-custom');

// Quick $5 button
if (quickBtn) {
  quickBtn.addEventListener('click', () => {
    selectedAmount = "5";
    quickBtn.classList.add('active');
    if (customInput) customInput.value = "";
  });
}

// Custom amount input
if (customInput) {
  customInput.addEventListener('input', () => {
    const val = customInput.value.trim();
    if (val && !isNaN(val) && Number(val) > 0) {
      selectedAmount = val;
      quickBtn.classList.remove('active');
    }
  });
}

if (supportBtn) {
  supportBtn.addEventListener('click', (e) => {
    e.stopPropagation();

    // Toggle popup
    paypalPopup.style.display = paypalPopup.style.display === 'block' ? 'none' : 'block';

    // GA4 tracking
    if (typeof gtag === "function") {
      gtag('event', 'support_click', {
        'event_category': 'Button',
        'event_label': 'Support Me'
      });
    }

    // Render PayPal once
    if (!isPayPalRendered && typeof paypal !== "undefined") {
      paypal.Buttons({
        style: { layout: 'vertical', color: 'gold', shape: 'rect', label: 'paypal', height: 40 },
        createOrder: (data, actions) => actions.order.create({
          purchase_units: [{ amount: { value: selectedAmount }, description: 'Support Payment' }]
        }),
        onApprove: (data, actions) => actions.order.capture().then(details => {
          alert('Thank you for your support, ' + details.payer.name.given_name);
          paypalPopup.style.display = 'none';
        }),
        onError: (err) => {
          console.error(err);
          alert('Payment could not be processed. Try again.');
        }
      }).render('#paypal-button-small');

      isPayPalRendered = true;
    }

    // Mobile positioning
    if (window.innerWidth <= 720) {
      paypalPopup.style.position = 'fixed';
      paypalPopup.style.bottom = '20px';
      paypalPopup.style.right = '20px';
      paypalPopup.style.top = 'auto';
      paypalPopup.style.width = '90%';
      paypalPopup.style.maxWidth = '320px';
    } else {
      paypalPopup.style.position = 'absolute';
      paypalPopup.style.top = '45px';
      paypalPopup.style.right = '0';
      paypalPopup.style.width = '300px';
    }
  });

  // Close popup if clicked outside
  window.addEventListener('click', (e) => {
    if (!e.target.closest('#supportBtn') && !e.target.closest('#paypalPopup')) {
      paypalPopup.style.display = 'none';
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
    emailjs.init('DhW4bXmuP0VP2d8bF');
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

      const templateParams = {
        software: feedbackForm.software.value || "Not selected",
        name: feedbackForm.name.value || "Anonymous",
        email: feedbackForm.email.value || "Not provided",
        message: feedbackForm.message.value
      };

      emailjs.send('zamini_musafir', 'template_yz15x2d', templateParams)
        .then(response => {
          statusEl.textContent = "Feedback sent successfully! Thank you.";
          feedbackForm.reset();
        })
        .catch(err => {
          console.error("EmailJS error:", err);
          statusEl.textContent = "Oops! Something went wrong. Check console.";
        });
    });
  }
  // ===== 7. Updates Subscription Form =====
  const updatesForm = document.getElementById('updatesForm');
  if (updatesForm) {
    const updatesStatus = document.getElementById('updates-status');
  
    updatesForm.addEventListener('submit', (e) => {
      e.preventDefault();
      updatesStatus.textContent = "Subscribing...";
  
      const templateParams = {
        email: updatesForm.email.value
      };
  
      emailjs.send('zamini_musafir', 'template_updates', templateParams)
        .then(() => {
          updatesStatus.textContent = "Subscribed successfully! You'll get updates soon.";
          updatesForm.reset();
        })
        .catch(err => {
          console.error("EmailJS error:", err);
          updatesStatus.textContent = "Oops! Something went wrong. Try again.";
        });
    });
  }




  
  // ===== 6. Track Download Buttons =====
  const downloadBtns = Array.from(document.querySelectorAll('a.btn'))
    .filter(btn => btn.textContent.trim() === 'Download');

  downloadBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      if (typeof gtag === "function") {
        gtag('event', 'download_click', {
          'event_category': 'Button',
          'event_label': 'Download'
        });
      }
    });
  });

});
gtag('event', 'subscribe_click', { 'event_category': 'Button', 'event_label': 'Updates' });
