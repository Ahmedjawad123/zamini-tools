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
// Reviews & Comment Approval
// ----------------------------
const submitBtn = document.getElementById('submit-review');
const reviewList = document.getElementById('review-list');
const reviewName = document.getElementById('review-name');
const reviewLocation = document.getElementById('review-location');
const reviewText = document.getElementById('review-text');
let selectedRating = 5; // default rating, can integrate star selection

// Array to store comments
let comments = [];

// Submit comment
submitBtn.addEventListener('click', () => {
  const name = reviewName.value.trim();
  const location = reviewLocation.value.trim();
  const text = reviewText.value.trim();

  if (!name || !text) {
    alert('Please enter your name/email and review text.');
    return;
  }

  // Add comment as pending
  comments.push({
    name: name,
    location: location,
    text: text,
    rating: selectedRating,
    approved: false
  });

  reviewName.value = '';
  reviewLocation.value = '';
  reviewText.value = '';

  displayComments();
});

// Display comments (approved first, pending after)
function displayComments() {
  reviewList.innerHTML = '';

  comments.forEach((c, index) => {
    const li = document.createElement('li');
    li.classList.add('review-item');
    li.style.padding = "8px";
    li.style.borderRadius = "5px";
    li.style.marginBottom = "8px";

    if (c.approved) {
      li.style.backgroundColor = "#1e1e1e";
      li.style.color = "#f0f0f0";
      li.innerHTML = `
        <div class="review-header">
          <b>${c.name}${c.location ? ', ' + c.location : ''}</b>
          <span class="review-stars">${'★'.repeat(c.rating)}</span>
        </div>
        <div class="review-text">"${c.text}"</div>
      `;
    } else {
      li.style.backgroundColor = "#333";
      li.style.color = "#fff";
      li.innerHTML = `
        <b>${c.name}${c.location ? ', ' + c.location : ''}</b> - Pending Approval
        <button class="approve-btn" data-index="${index}">Approve</button>
        <button class="delete-btn" data-index="${index}">Delete</button>
      `;
    }

    reviewList.appendChild(li);
  });

  // Attach approve/delete handlers for pending comments
  document.querySelectorAll('.approve-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const i = btn.dataset.index;
      comments[i].approved = true;
      displayComments();
    });
  });

  document.querySelectorAll('.delete-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const i = btn.dataset.index;
      comments.splice(i, 1);
      displayComments();
    });
  });
}

// Initialize display
displayComments();

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

// Close PayPal popup on outside click
window.addEventListener('click', (e) => {
  if (!e.target.closest('#supportBtn') && !e.target.closest('#paypalPopup')) {
    paypalPopup.style.display = 'none';
  }
});
