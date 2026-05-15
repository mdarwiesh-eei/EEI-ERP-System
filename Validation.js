// Validate customer name (Windows safe)
function isValidName(name) {
  return !(/[\\\/:*?"<>|]/.test(name));
}

function isValidPhone(phone) {

  if (!phone) return true;

  return /^[0-9]+$/.test(phone);
}

function isValidEmail(email) {

  if (!email) return true;

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}




