// ================== main.js ==================

document.addEventListener('DOMContentLoaded', () => {

  // ===== 1. Footer Year =====
  const yearEl = document.getElementById("year");
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  // ===== 2. Smooth Scroll =====
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function(e) {
      e.preventDefault();
      const target = document.querySelector(this.getAttribute('href'));
      if (target) target.scrollIntoView({ behavior: 'smooth' });
    });
  });

  // ===== 3. Tabs Functionality =====
  const tabs = document.querySelectorAll('.tab');
  const contents = document.querySelectorAll('.tab-content');
  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      const target = document.getElementById(tab.dataset.tab);
      contents.forEach(c => c.classList.remove('active'));
      if (target) target.classList.add('active');
    });
  });

  // ===== 4. Views Counter (Client-side only) =====
  const totalViewsEl = document.getElementById('totalViews');
  if (totalViewsEl) {
    let views = parseInt(localStorage.getItem('totalViews')) || 0;
    views += 1;
    localStorage.setItem('totalViews', views);
    totalViewsEl.textContent = views;
  }

  // ===== 5. Chat Toggle =====
  const chatBtn = document.getElementById('chatBtn');
  const chatBox = document.getElementById('chatBox');
  if (chatBtn && chatBox) {
    chatBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      chatBox.style.display = chatBox.style.display === 'block' ? 'none' : 'block';
    });

    // Close chat when clicking outside
    window.addEventListener('click', (e) => {
      if (!e.target.closest('#chatBtn') && !e.target.closest('#chatBox')) {
        chatBox.style.display = 'none';
      }
    });
  }

  // ===== 6. Support Button & PayPal =====
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

  if (supportBtn && paypalPopup) {
    supportBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      paypalPopup.style.display = paypalPopup.style.display === 'block' ? 'none' : 'block';

      if (!isPayPalRendered && typeof paypal !== "undefined") {
        paypal.Buttons({
          style: { layout: 'vertical', color: 'gold', shape: 'rect', label: 'paypal', height: 40 },
          createOrder: (data, actions) => actions.order.create({
            purchase_units: [{ amount: { value: selectedAmount }, description: 'Support Payment' }]
          }),
          onApprove: (data, actions) => actions.order.capture().then(details => {
            alert('Thank you for your support, ' + details.payer.name.given_name + '!');
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
        paypalPopup.style.width = '90%';
        paypalPopup.style.maxWidth = '320px';
      } else {
        paypalPopup.style.position = 'absolute';
        paypalPopup.style.top = '45px';
        paypalPopup.style.right = '0';
        paypalPopup.style.width = '300px';
      }
    });

    window.addEventListener('click', (e) => {
      if (!e.target.closest('#supportBtn') && !e.target.closest('#paypalPopup')) {
        paypalPopup.style.display = 'none';
      }
    });
  }

  // ===== 7. Download Buttons Tracking =====
  const downloadBtns = document.querySelectorAll('a.btn-download');
  downloadBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      if (typeof gtag === "function") {
        gtag('event', 'download_click', {
          'event_category': 'Button',
          'event_label': btn.textContent.trim()
        });
      }
    });
  });

  // ===== 8. File Size Fetch =====
  const fileSizeSpan = document.getElementById('file-size');
  const fileUrl = "https://github.com/Ahmedjawad123/Zamini_Converter/releases/download/v1.0.0/Executable_file_.Zamini_Converter_v1.0.0.rar";
  if (fileSizeSpan) {
    fetch(fileUrl, { method: 'HEAD' })
      .then(resp => {
        const size = resp.headers.get('content-length');
        fileSizeSpan.textContent = size ? (size / (1024 * 1024)).toFixed(2) + " MB" : "N/A";
      })
      .catch(err => {
        console.error('File size fetch error:', err);
        fileSizeSpan.textContent = "N/A";
      });
  }

  // ===== 9. Initialize EmailJS =====
  if (typeof emailjs !== "undefined") emailjs.init('DhW4bXmuP0VP2d8bF');

  // ===== 10. Feedback Form =====
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
        software: feedbackForm.software?.value || "Not selected",
        name: feedbackForm.name?.value || "Anonymous",
        email: feedbackForm.email?.value || "Not provided",
        message: feedbackForm.message?.value || "No message"
      };

      emailjs.send('zamini_musafir', 'template_yz15x2d', templateParams)
        .then(() => {
          statusEl.textContent = "Feedback sent successfully! Thank you.";
          feedbackForm.reset();
        })
        .catch(err => {
          console.error("EmailJS error:", err);
          statusEl.textContent = "Oops! Something went wrong. Check console.";
        });
    });
  }

  // ===== 11. Updates Subscription Form =====
  const updatesForm = document.getElementById('updatesForm');
  if (updatesForm) {
    const updatesStatus = document.getElementById('updates-status');

    updatesForm.addEventListener('submit', (e) => {
      e.preventDefault();
      updatesStatus.textContent = "Subscribing...";

      const email = updatesForm.email.value.trim();
      if (!email) {
        updatesStatus.textContent = "Enter a valid email!";
        return;
      }

      if (typeof emailjs !== "undefined") {
        emailjs.send('zamini_musafir', 'template_updates', { email })
          .then(() => {
            updatesStatus.textContent = "Subscribed successfully! You'll get updates soon.";
            updatesForm.reset();
          })
          .catch(err => {
            console.error("EmailJS error:", err);
            updatesStatus.textContent = "Oops! Something went wrong. Try again.";
          });
      } else {
        console.error("EmailJS not loaded!");
        updatesStatus.textContent = "Email service not available.";
      }
    });
  }



  // ===== Download Counter =====
const downloadWrappers = document.querySelectorAll('.download-wrapper');

downloadWrappers.forEach((wrapper, index) => {
  const btn = wrapper.querySelector('.btn-download');
  const countEl = wrapper.querySelector('.count');
  const storageKey = `download_count_${index}`;

  // Load initial count from localStorage
  let count = parseInt(localStorage.getItem(storageKey)) || 0;
  countEl.textContent = count;

  // Increment on click
  btn.addEventListener('click', () => {
    count += 1;
    localStorage.setItem(storageKey, count);
    countEl.textContent = count;

    // Optional: Google Analytics tracking
    if (typeof gtag === "function") {
      gtag('event', 'download_click', {
        'event_category': 'Button',
        'event_label': btn.textContent.trim()
      });
    }
  });
});


});
