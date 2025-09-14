// Set current year in footer
document.getElementById("year").textContent = new Date().getFullYear();

// Main image change with caption
const mainImg = document.getElementById("main-img");
const captionText = document.getElementById("caption-text"); // optional if caption element exists
const thumbs = document.querySelectorAll(".thumbs img");

thumbs.forEach(img => {
  img.addEventListener("click", () => {
    mainImg.src = img.src;
    if(captionText) captionText.textContent = img.dataset.caption;
    thumbs.forEach(t => t.classList.remove("active"));
    img.classList.add("active");
  });
});

// Double-click zoom
mainImg.addEventListener("dblclick", () => { 
  window.open(mainImg.src, "_blank"); 
});

// Star rating
const stars = document.querySelectorAll(".stars span");
let selectedRating = 0;
stars.forEach(star => {
  star.addEventListener("click", () => {
    selectedRating = star.dataset.value;
    stars.forEach(s => s.classList.remove("selected"));
    for(let i=0;i<selectedRating;i++){ stars[i].classList.add("selected"); }
  });
});

// Reviews
const submitReview = document.getElementById("submit-review");
const reviewList = document.getElementById("review-list");

submitReview.addEventListener("click", () => {
  const name = document.getElementById("review-name").value.trim();
  const location = document.getElementById("review-location").value.trim();
  const text = document.getElementById("review-text").value.trim();

  if(!name || !text){ 
    alert("Name/Email and Review text are required."); 
    return; 
  }

  const li = document.createElement("li");
  li.innerHTML = `<strong>${name}</strong>${location ? ' (' + location + ')' : ''} <br>
                  <em>${new Date().toLocaleString()}</em> <br>
                  ${text} <br>
                  <strong>Rating:</strong> ${selectedRating || 'N/A'} ★`;
  reviewList.appendChild(li);

  document.getElementById("review-name").value = "";
  document.getElementById("review-location").value = "";
  document.getElementById("review-text").value = "";
  selectedRating = 0;
  stars.forEach(s => s.classList.remove("selected"));
});

// PayPal Support Button
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
