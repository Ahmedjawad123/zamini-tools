// ===== Firebase Initialization =====
if (typeof firebase === "undefined") {
  console.error("Firebase SDK not loaded!");
} else {
  const firebaseConfig = {
    apiKey: "AIzaSyDUUMyJDZXdGa1LyxcESOcth3e3ZPovt-0",
    authDomain: "zaminimusafir.firebaseapp.com",
    projectId: "zaminimusafir",
    storageBucket: "zaminimusafir.firebasestorage.app",
    messagingSenderId: "1066132693199",
    appId: "1:1066132693199:web:8b87e2c3270434891d17ba",
    measurementId: "G-YVCFZ783GR"
  };

  if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
    if (firebase.analytics) firebase.analytics();
  }

  const db = firebase.firestore();

  // ===== Ensure document exists =====
  async function ensureDoc(productName) {
    const docRef = db.collection("downloads").doc(productName);
    let doc = await docRef.get();
    if (!doc.exists) {
      await docRef.set({ count: 0 });
      doc = await docRef.get(); // re-fetch after creation
    }
    return docRef;
  }

  // ===== Increment download count =====
  async function incrementDownload(productName) {
    const docRef = db.collection("downloads").doc(productName);
    await docRef.set({ count: firebase.firestore.FieldValue.increment(1) }, { merge: true });
  }

  // ===== Initialize all download buttons =====
  document.querySelectorAll('a.btn[data-product]').forEach(async btn => {
    const productName = btn.dataset.product;
    const countEl = document.querySelector(`.download-count[data-product="${productName}"]`);

    // Ensure document exists
    const docRef = await ensureDoc(productName);

    // Show current count on page load
    const doc = await docRef.get();
    if (countEl && doc.data()) countEl.textContent = doc.data().count;

    // Increment on click
    btn.addEventListener('click', async () => {
      await incrementDownload(productName);
    });

    // Real-time updates
    docRef.onSnapshot(doc => {
      if (countEl && doc.exists && doc.data()) {
        countEl.textContent = doc.data().count;
      }
    });
  });
}



document.addEventListener('DOMContentLoaded', () => {

  // ===== 1. Footer Year =====
  const yearEl = document.getElementById("year");
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  // ===== 2. Support Button & PayPal =====
  const supportBtn = document.getElementById('supportBtn');
  const paypalPopup = document.getElementById('paypalPopup');
  let isPayPalRendered = false;
  let selectedAmount = "5"; // default

  const quickBtn = document.getElementById('donate-5');
  const customInput = document.getElementById('donate-custom');

  if (quickBtn) quickBtn.addEventListener('click', () => {
    selectedAmount = "5";
    quickBtn.classList.add('active');
    if (customInput) customInput.value = "";
  });

  if (customInput) customInput.addEventListener('input', () => {
    const val = customInput.value.trim();
    if (val && !isNaN(val) && Number(val) > 0) {
      selectedAmount = val;
      quickBtn.classList.remove('active');
    }
  });

  if (supportBtn) {
    supportBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      paypalPopup.style.display = paypalPopup.style.display === 'block' ? 'none' : 'block';

      // GA4 tracking
      if (typeof gtag === "function") {
        gtag('event', 'support_click', {
          'event_category': 'Button',
          'event_label': 'Support Me'
        });
      }

      // Render PayPal button once
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
        fileSizeSpan.textContent = size ? (size / (1024 * 1024)).toFixed(2) + " MB" : "N/A";
      })
      .catch(err => {
        console.error('File size fetch error:', err);
        fileSizeSpan.textContent = "N/A";
      });
  }

  // ===== 4. Initialize EmailJS =====
  if (typeof emailjs !== "undefined") emailjs.init('DhW4bXmuP0VP2d8bF');

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

  // ===== 6. Updates Form =====
  const updatesForm = document.getElementById('updatesForm');
  if (updatesForm) {
    const updatesStatus = document.getElementById('updates-status');
    updatesForm.addEventListener('submit', (e) => {
      e.preventDefault();
      updatesStatus.textContent = "Subscribing...";
      const templateParams = { email: updatesForm.email.value };

      emailjs.send('zamini_musafir', 'template_updates', templateParams)
        .then(() => {
          updatesStatus.textContent = "Subscribed successfully!";
          updatesForm.reset();
        })
        .catch(err => {
          console.error("EmailJS error:", err);
          updatesStatus.textContent = "Oops! Something went wrong. Try again.";
        });
    });
  }

  // ===== 7. Track Download Buttons =====
  document.querySelectorAll('a.btn').forEach(btn => {
    if (btn.textContent.trim().includes('Download')) {
      btn.addEventListener('click', () => {
        if (typeof gtag === "function") {
          gtag('event', 'download_click', { 'event_category': 'Button', 'event_label': 'Download' });
        }
      });
    }
  });

  // ===== 8. GA4 Subscribe Click Event (Optional) =====
  if (typeof gtag === "function") {
    gtag('event', 'subscribe_click', { 'event_category': 'Button', 'event_label': 'Updates' });
  }

});
