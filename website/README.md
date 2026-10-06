# Saxon I-REC portal (website)

The website for the Saxon Renewables I-REC portal, published with GitHub Pages.

This repository is public on purpose and contains **no secrets**. The page signs staff in with their Microsoft work account and reads and writes the portal's SharePoint lists as that person. It never contacts the I-REC registry; that is done by the private worker.

- `frontend/config.js`: tenant ID, app ID and SharePoint location. Identifiers only.
- `frontend/vendor/`: Microsoft's sign-in library (MSAL 3.30.0), bundled so nothing loads from elsewhere.
- Publishing refuses to run if anything in `frontend` looks like a secret.

Setup: see SETUP.md in the worker repository.
