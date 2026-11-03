// Sign-in settings for this page: the Entra app "Shalam Tool" (single tenant). A client ID is public, not a secret;
// if it is ever blank the page says so and stops. Sources: where each file or folder is, as a SharePoint link or a
// /sites/<site>/<library>/<path>; a blank one shows grey "not set". The error email goes to errorEmailTo.
// The sales sheet (Saul, Nov 2): saved as salesFile into salesLibrary (a /sites/<site>/<library>), the one place the page
// may write; blank turns saving off. Its photos come from Quote Tracker's store: <photos>/thumb/<KEY>.jpg.
window.INV_CONFIG = {
  tenantId: 'eeea7e3e-9ccb-4a16-894e-3417c6a7a4c1',
  clientId: 'f994e0f2-f0d4-452e-88d6-67d9396d7552',
  sources: {
    msl: 'https://met1sol.sharepoint.com/:x:/s/Production/IQBIuuE0Jan6RaTDWpWOz8W-AZzN3lzg4rcL1r0SraBDZ_c?e=hshN16',
    tracker: 'https://met1sol.sharepoint.com/:x:/s/Production/IQA76OYghhjcS6oOzo2fIg3HAfPnwKQxilgLxbS6m2WPClM?e=q9bZXq',
    ct: '/sites/Importing/Shared Documents/Importing/Logistics/Container Tracking Sheet.xlsx',
    forwarder: '/sites/Reports/Shared Documents/Reports/Shipping Reports/Forwarder',
    drayage: '/sites/Reports/Shared Documents/Reports/Shipping Reports/Drayage',
    ats: '/sites/Reports/Shared Documents/Reports/Shipping Reports/ATS',
  },
  // Container Tracking sections with no round: straight to the customer (Pull from China), not booked yet (Master Sheet stage), or ignored.
  sections: { pullFromChina: ['GT', 'Giant Tiger', 'TJX Mexico', 'TJX Aus', 'TJX UK', 'FOB Loaded'], planning: ['Unshipped', 'Planning - NB', 'Planning - other'], ignore: ['CT - Hidden Rows'] },
  salesLibrary: '/sites/Reports/Shalam Sales tool',
  salesFile: 'Available to sell.xlsx',
  photos: '/sites/Production/Shared Documents/Quote Sheets/Quote Images',
  errorEmailTo: 'saul@metro1solutions.com',
  errorEmailName: 'Saul',
};
