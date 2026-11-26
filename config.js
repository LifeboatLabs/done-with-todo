/*
 * Done With To Do: settings.
 *
 * Paste your values between the quotes, then click "Commit changes".
 * The setup guide says where each one comes from.
 */
window.DWTD_CONFIG = {

  // Microsoft Entra -> App registrations -> Done With To Do -> "Application (client) ID"
  microsoftClientId: "8f748ea4-dc22-4e4d-a90d-c68aca9c4798",

  // Google Cloud -> Google Auth Platform -> Clients -> your web client -> "Client ID"
  // Leave as it is to launch without Google Tasks; the button will say "Coming soon".
  googleClientId: "49981952788-flgm821vvrnti82br3ska88jse680e9n.apps.googleusercontent.com",

  // Keep false until Google approves the app. While false, only testers see the Google Tasks
  // button: open the site with ?testers on the end. Then change it to true for everyone.
  googleForEveryone: false,

  // Your tip or sponsor page, for example "https://ko-fi.com/yourname"
  // or "https://github.com/sponsors/yourname". Leave empty to hide the tip button.
  donateUrl: "https://lifeboatlabs.github.io/tip/",

  // Shown on the privacy page. Google asks for a way to contact the app's owner.
  ownerName: "Lifeboat Labs",
  contactEmail: "lifeboatlabs@outlook.com",

  // Leave these alone unless asked to change them.
  githubUrl: "",              // worked out automatically on GitHub Pages
  microsoftAuthority: "common"
};
