// ----------------------------
// Footer Year
// ----------------------------
document.getElementById("year").textContent = new Date().getFullYear();

// ----------------------------
// Main Image & Thumbnails
// ----------------------------
const mainImg = document.getElementById("main-img");
const thumbs = document.querySelectorAll(".thumbs img");

thumbs.forEach(img => {
  img.addEventListener("click", () => {
    mainImg.src = img.src;
    thumbs.forEach(t => t.classList.remove("active"));
    img.classList.add("active");
  });
});

// Double-click zoom
mainImg.addEventListener("dblclick", () => {
  window.open(mainImg.src, "_blank");
});

// ----------------------------
// Reviews
// ----------------------------
const submitBtn = document.getElementById('submit-review');
const reviewList = document.getElementById('review-list');
const reviewName = document.getElementById('review-name');
const reviewLocation = document.getElementById('review-location');
const reviewText = document.getElementById('review-text');

submitBtn.addEventListener('click', function() {
  const name = reviewName.value.trim();
  const location = reviewLocation.value.trim();
  const text = reviewText.value.trim();

  if (!name || !text) {
    alert('Please enter your name/email and review text.');
    return;
  }

  // Create new review item
  const li = document.createElement('li');
  li.classList.add('review-item'); // CSS handles background, padding, margin, color


  const header = document.createElement('div');
  header.classList.add('review-header');
  header.innerHTML = `<b>${name}${location ? ', ' + location : ''}</b> <span class="review-stars">★★★★★</span>`;

  const reviewDiv = document.createElement('div');
  reviewDiv.classList.add('review-text');
  reviewDiv.innerText = `"${text}"`;

  li.appendChild(header);
  li.appendChild(reviewDiv);

  // Prepend new review at top
  reviewList.insertBefore(li, reviewList.firstChild);

  // Clear input fields
  reviewName.value = '';
  reviewLocation.value = '';
  reviewText.value = '';
});

// ----------------------------
// PayPal Support Button
// ----------------------------
const supportBtn = document.getElementById('supportBtn');
const paypalPopup = document.getElementById('paypalPopup');
let isPayPalRendered = false;

supportBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  paypalPopup.style.display = paypalPopup.style.display === 'block' ? 'none' : 'block';

  if (!isPayPalRendered && typeof paypal !== "undefined") {
    paypal.Buttons({
      style: { layout: 'vertical', color: 'gold', shape: 'rect', label: 'paypal', height: 40 },
      createOrder: function(data, actions) {
        return actions.order.create({
          purchase_units: [{ amount: { value: '5' }, description: 'Support Payment' }]
        });
      },
      onApprove: function(data, actions) {
        return actions.order.capture().then(function(details) {
          alert('Thank you for your support, ' + details.payer.name.given_name);
          paypalPopup.style.display = 'none';
        });
      },
      onError: function(err) {
        console.error(err);
        alert('Payment could not be processed. Try again.');
      }
    }).render('#paypal-button-small');

    isPayPalRendered = true;
  }
});

// Close popup when clicking outside
window.addEventListener('click', (e) => {
  if (!e.target.closest('#supportBtn') && !e.target.closest('#paypalPopup')) {
    paypalPopup.style.display = 'none';
  }
});

