// ===== 0. Firebase Initialization =====
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

  // Only initialize once
  if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);

    if (firebase.analytics) {
      firebase.analytics();
      console.log("Firebase Analytics initialized.");
    } else {
      console.warn("Firebase Analytics not loaded — heartbeats disabled.");
    }
  }

  // ✅ Initialize db once, outside the if block
  const db = firebase.firestore();

  // ===== Increment and track download counts =====
  function incrementDownload(productName) {
    const docRef = db.collection("downloads").doc(productName);
    docRef.set({ count: firebase.firestore.FieldValue.increment(1) }, { merge: true })
      .catch(err => console.error("Failed to increment download:", err));
  }

  // ===== Real-time listener for download counts =====
  function listenDownloadCount(productName) {
    const docRef = db.collection("downloads").doc(productName);
    const countEl = document.querySelector(`.download-count[data-product="${productName}"]`);

    if (countEl) {
      docRef.onSnapshot(doc => {
        if (doc.exists) {
          countEl.textContent = doc.data().count || 0;
        } else {
          countEl.textContent = 0;
        }
      });
    }
  }

  // ===== Initialize for all products =====
  document.querySelectorAll('a.btn[data-product]').forEach(btn => {
    const productName = btn.dataset.product;
    btn.addEventListener('click', () => incrementDownload(productName));
    listenDownloadCount(productName);
  });
}












// ===== DOMContentLoaded =====
document.addEventListener('DOMContentLoaded', () => {

  // ===== Footer Year =====
  const yearEl = document.getElementById("year");
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  // ===== PayPal Support Button =====
  const supportBtn = document.getElementById('supportBtn');
  const paypalPopup = document.getElementById('paypalPopup');
  let isPayPalRendered = false;
  let selectedAmount = "5";
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

  // ===== Feedback Form =====
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

  // ===== Updates Form =====
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

  // ===== Track Download Buttons with GA4 =====
  document.querySelectorAll('a.btn[data-product]').forEach(btn => {
    btn.addEventListener('click', () => {
      if (typeof gtag === "function") {
        gtag('event', 'download_click', { 
          event_category: 'Button', 
          event_label: btn.dataset.product 
        });
      }
    });
  });

});
