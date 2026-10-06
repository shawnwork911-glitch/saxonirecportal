# Saxon I-REC portal (website)

The website for the Saxon Renewables I-REC portal, published with GitHub Pages.

This repository is public on purpose and contains **no secrets**. Staff sign in with their Microsoft work account. The page reads and writes the portal's SharePoint lists as that person, and logs every action. It never contacts the I-REC registry: registry data is exported by the private exporter repository and imported here.

- `frontend/config.js`: tenant ID, app ID and SharePoint location. Identifiers only.
- `frontend/vendor/`: Microsoft's sign-in library (MSAL 3.30.0), bundled.
- Publishing refuses to run if anything in `frontend` looks like a secret.
