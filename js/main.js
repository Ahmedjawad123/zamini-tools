document.addEventListener('DOMContentLoaded', () => {

  // ===== 1. Footer Year =====
  const yearEl = document.getElementById("year");
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  // ===== 2. Smooth Scroll =====
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
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

  // ===== 4. Firebase Views Counter =====
  if (typeof firebase !== "undefined") {
    const firebaseConfig = {
      apiKey: "AIzaSyDUUMyJDZXdGa1LyxcESOcth3e3ZPovt-0",
      authDomain: "zaminimusafir.firebaseapp.com",
      projectId: "zaminimusafir",
      storageBucket: "zaminimusafir.firebasestorage.app",
      messagingSenderId: "1066132693199",
      appId: "1:1066132693199:web:8b87e2c3270434891d17ba",
      measurementId: "G-YVCFZ783GR"
    };
    if (!firebase.apps.length) firebase.initializeApp(firebaseConfig);
    const db = firebase.firestore();
    const totalViewsEl = document.getElementById('totalViews');

    if (totalViewsEl) {
      const viewsRef = db.collection("siteStats").doc("totalViews");

      viewsRef.update({ count: firebase.firestore.FieldValue.increment(1) })
        .catch(() => viewsRef.set({ count: 1 }));

      viewsRef.onSnapshot(doc => {
        totalViewsEl.textContent = doc.exists ? doc.data().count || 0 : 0;
      });
    }
  }

  // ===== 5. Chat Toggle =====
  const chatBtn = document.getElementById('chatBtn');
  const chatBox = document.getElementById('chatBox');
  if (chatBtn && chatBox) {
    chatBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      chatBox.style.display = chatBox.style.display === 'block' ? 'none' : 'block';
    });
    window.addEventListener('click', (e) => {
      if (!e.target.closest('#chatBtn') && !e.target.closest('#chatBox')) {
        chatBox.style.display = 'none';
      }
    });
  }

  // ===== 6. Support / PayPal Popup =====
  const supportBtn = document.getElementById('supportBtn');
  const paypalPopup = document.getElementById('paypalPopup');
  const quickBtn = document.getElementById('donate-5');
  const customInput = document.getElementById('donate-custom');
  let isPayPalRendered = false;

  if (supportBtn && paypalPopup) {
    if (quickBtn) {
      quickBtn.addEventListener('click', () => {
        quickBtn.classList.add('active');
        if (customInput) customInput.value = "";
      });
    }
    if (customInput) {
      customInput.addEventListener('input', () => { if (quickBtn) quickBtn.classList.remove('active'); });
    }

    supportBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      paypalPopup.style.display = paypalPopup.style.display === 'block' ? 'none' : 'block';

      if (!isPayPalRendered && typeof paypal !== "undefined") {
        paypal.Buttons({
          style: { layout: 'vertical', color: 'gold', shape: 'rect', label: 'paypal', height: 40 },
          createOrder: (data, actions) => {
            const amount = customInput && customInput.value && !isNaN(customInput.value) && Number(customInput.value) > 0 ? customInput.value : "5";
            return actions.order.create({ purchase_units: [{ amount: { value: amount }, description: 'Support Payment' }] });
          },
          onApprove: (data, actions) => actions.order.capture().then(details => {
            alert('Thank you for your support, ' + details.payer.name.given_name + '!');
            paypalPopup.style.display = 'none';
          }),
          onError: (err) => { console.error(err); alert('Payment could not be processed.'); }
        }).render('#paypal-button-small');
        isPayPalRendered = true;
      }

      // Position popup
      const posPopup = () => {
        if (window.innerWidth <= 720) {
          paypalPopup.style.position = 'fixed';
          paypalPopup.style.left = '5%';
          paypalPopup.style.bottom = '20px';
          paypalPopup.style.width = '90%';
          paypalPopup.style.maxWidth = '320px';
        } else {
          paypalPopup.style.position = 'absolute';
          paypalPopup.style.top = `${supportBtn.offsetTop + supportBtn.offsetHeight + 8}px`;
          paypalPopup.style.left = `${supportBtn.offsetLeft}px`;
          paypalPopup.style.width = '300px';
        }
      };
      posPopup();
      window.addEventListener('resize', posPopup);
      window.addEventListener('scroll', posPopup);
    });

    window.addEventListener('click', (e) => {
      if (!e.target.closest('#supportBtn') && !e.target.closest('#paypalPopup')) paypalPopup.style.display = 'none';
    });
  }

  // ===== 7. Download Buttons Tracking =====
  document.querySelectorAll('a.btn-download').forEach((btn, index) => {
    const storageKey = `download_count_${index}`;
    let countEl = btn.nextElementSibling;

    if (!countEl || !countEl.classList.contains('count')) {
      countEl = document.createElement('span');
      countEl.className = 'count';
      countEl.style.marginLeft = "10px";
      btn.insertAdjacentElement('afterend', countEl);
    }

    let count = parseInt(localStorage.getItem(storageKey)) || 0;
    countEl.textContent = count;

    btn.addEventListener('click', () => {
      count++;
      localStorage.setItem(storageKey, count);
      countEl.textContent = count;
    });
  });

  // ===== 8. EmailJS Init =====
  if (typeof emailjs !== "undefined") emailjs.init('DhW4bXmuP0VP2d8bF');

  // ===== 9. Feedback Form =====
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
        .then(() => { statusEl.textContent = "Feedback sent successfully!"; feedbackForm.reset(); })
        .catch(err => { console.error(err); statusEl.textContent = "Oops! Something went wrong."; });
    });
  }

  // ===== 10. Updates Subscription Form =====
  const updatesForm = document.getElementById('updatesForm');
  if (updatesForm) {
    const updatesStatus = document.getElementById('updates-status');
    updatesForm.addEventListener('submit', (e) => {
      e.preventDefault();
      updatesStatus.textContent = "Subscribing...";
      const email = updatesForm.email.value.trim();
      if (!email) { updatesStatus.textContent = "Enter a valid email!"; return; }
      if (typeof emailjs !== "undefined") {
        emailjs.send('zamini_musafir', 'template_updates', { email })
          .then(() => { updatesStatus.textContent = "Subscribed successfully!"; updatesForm.reset(); })
          .catch(err => { console.error(err); updatesStatus.textContent = "Oops! Something went wrong."; });
      } else updatesStatus.textContent = "Email service not available.";
    });
  }

  // ===== 11. FAQ Collapse/Expand =====
  const faqQuestions = document.querySelectorAll('.faq-question');
  faqQuestions.forEach(question => {
    question.addEventListener('click', () => {
      const answer = question.nextElementSibling;
      answer.classList.toggle('open');
    });
  });






  
/* ===== Tab Visibility Control ===== */
.tab-content {
  display: none;
  padding: 20px;
  animation: fadeIn 0.3s ease-in-out;
}
.tab-content.active {
  display: block;
}

@keyframes fadeIn {
  from { opacity: 0; transform: translateY(10px); }
  to { opacity: 1; transform: translateY(0); }
}

});
