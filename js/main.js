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

  // ===== 4. Views Counter (Persistent) =====
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
    window.addEventListener('click', (e) => {
      if (!e.target.closest('#chatBtn') && !e.target.closest('#chatBox')) {
        chatBox.style.display = 'none';
      }
    });
  }

  // ===== 6. Support Button & PayPal Popup =====
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
      customInput.addEventListener('input', () => {
        if (quickBtn) quickBtn.classList.remove('active');
      });
    }

    supportBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      paypalPopup.style.display = paypalPopup.style.display === 'block' ? 'none' : 'block';

      // Render PayPal button only once
      if (!isPayPalRendered && typeof paypal !== "undefined") {
        paypal.Buttons({
          style: {
            layout: 'vertical',
            color: 'gold',
            shape: 'rect',
            label: 'paypal',
            height: 40
          },
          createOrder: (data, actions) => {
            const amount =
              customInput && customInput.value && !isNaN(customInput.value) && Number(customInput.value) > 0
                ? customInput.value
                : "5";
            return actions.order.create({
              purchase_units: [{ amount: { value: amount }, description: 'Support Payment' }]
            });
          },
          onApprove: (data, actions) => {
            return actions.order.capture().then(details => {
              alert('Thank you for your support, ' + details.payer.name.given_name + '!');
              paypalPopup.style.display = 'none';
            });
          },
          onError: (err) => {
            console.error(err);
            alert('Payment could not be processed. Try again.');
          }
        }).render('#paypal-button-small');
        isPayPalRendered = true;
      }

      // ===== Position popup under button =====
      const btnRect = supportBtn.getBoundingClientRect();
      if (window.innerWidth <= 720) {
        // Mobile
        paypalPopup.style.position = 'fixed';
        paypalPopup.style.top = 'auto';
        paypalPopup.style.left = '5%';
        paypalPopup.style.bottom = '20px';
        paypalPopup.style.width = '90%';
        paypalPopup.style.maxWidth = '320px';
      } else {
        // Desktop
        paypalPopup.style.position = 'absolute';
        paypalPopup.style.top = `${supportBtn.offsetTop + supportBtn.offsetHeight + 8}px`;
        paypalPopup.style.left = `${supportBtn.offsetLeft}px`;
        paypalPopup.style.width = '300px';
      }
    });

    // Close popup when clicking outside
    window.addEventListener('click', (e) => {
      if (!e.target.closest('#supportBtn') && !e.target.closest('#paypalPopup')) {
        paypalPopup.style.display = 'none';
      }
    });

    // Adjust popup on resize
    window.addEventListener('resize', () => {
      if (paypalPopup.style.display === 'block') {
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
      }
    });

    // Adjust popup on scroll
    window.addEventListener('scroll', () => {
      if (paypalPopup.style.display === 'block' && window.innerWidth > 720) {
        paypalPopup.style.top = `${supportBtn.offsetTop + supportBtn.offsetHeight + 8}px`;
        paypalPopup.style.left = `${supportBtn.offsetLeft}px`;
      }
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

  // ===== 8. Initialize EmailJS =====
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

  // ===== 10. Updates Subscription Form =====
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
        updatesStatus.textContent = "Email service not available.";
      }
    });
  }

});
