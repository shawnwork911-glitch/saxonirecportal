// Saxon I-REC portal settings. These are identifiers, not secrets: it is safe for
// them to be public. Never put a password or the I-REC secret in this repository.
window.IREC_PORTAL_CONFIG = {
  tenantId: '<TENANT_ID>',                 // Microsoft Entra ID: Directory (tenant) ID
  clientId: '<PORTAL_CLIENT_ID>',          // "Saxon I-REC portal" app: Application (client) ID
  siteUrl: 'https://saxonrenewables.sharepoint.com/sites/Operations',
  library: 'IREC Portal',
  folder: 'Portal data'
};
