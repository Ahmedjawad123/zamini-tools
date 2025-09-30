document.addEventListener('DOMContentLoaded', () => {

  // ===== Footer Year =====
  const yearEl = document.getElementById("year");
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  // ===== Download buttons with Firebase =====
  const db = firebase.firestore();

  function incrementDownload(productName) {
    const docRef = db.collection("downloads").doc(productName);
    docRef.get().then(docSnap => {
      if (!docSnap.exists) {
        docRef.set({ count: 1 });
      } else {
        docRef.update({ count: firebase.firestore.FieldValue.increment(1) });
      }
    }).catch(err => console.error("Download count error:", err));
  }

  function listenDownloadCount(productName) {
    const countEl = document.querySelector(`.download-count[data-product="${productName}"]`);
    if (!countEl) return;
    db.collection("downloads").doc(productName)
      .onSnapshot(docSnap => {
        countEl.textContent = docSnap.exists ? (docSnap.data().count || 0) : 0;
      });
  }

  document.querySelectorAll('a.btn[data-product]').forEach(btn => {
    const productName = btn.dataset.product;
    btn.addEventListener('click', () => incrementDownload(productName));
    listenDownloadCount(productName);

    // GA4 tracking (optional)
    btn.addEventListener('click', () => {
      if (typeof gtag === "function") {
        gtag('event', 'download_click', { event_category: 'Button', event_label: productName });
      }
    });
  });












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
