/*
 * Done With To Do: settings.
 *
 * Paste your values between the quotes, then click "Commit changes".
 * The setup guide says where each one comes from.
 */
window.DWTD_CONFIG = {

  // Microsoft Entra -> App registrations -> Done With To Do -> "Application (client) ID"
  microsoftClientId: "PASTE-YOUR-MICROSOFT-CLIENT-ID-HERE",

  // Google Cloud -> Google Auth Platform -> Clients -> your web client -> "Client ID"
  // Leave as it is to launch without Google Tasks; the button will say "Coming soon".
  googleClientId: "PASTE-YOUR-GOOGLE-CLIENT-ID-HERE",

  // Your tip or sponsor page, for example "https://ko-fi.com/yourname"
  // or "https://github.com/sponsors/yourname". Leave empty to hide the tip button.
  donateUrl: "",

  // Shown on the privacy page. Google asks for a way to contact the app's owner.
  ownerName: "Lifeboat Labs",
  contactEmail: "lifeboatlabs@outlook.com",

  // Leave these alone unless asked to change them.
  githubUrl: "",              // worked out automatically on GitHub Pages
  microsoftAuthority: "common"
};
