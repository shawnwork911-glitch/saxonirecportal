// Saxon I-REC portal settings. These are identifiers, not secrets: it is safe for
// them to be public. Never put a password or the I-REC secret in this repository.
window.IREC_PORTAL_CONFIG = {
  tenantId: 'ab7c8c1e-dc1d-431f-ba8b-e4de2b047f8b',   // Saxon Renewable Energy: Directory (tenant) ID
  clientId: '1517c640-f9f8-429d-9471-c497ae2932ee',   // Saxon IREC Portal: Application (client) ID
  siteUrl: 'https://saxonrenewables.sharepoint.com/sites/Operations',
  library: 'IREC Portal',
  folder: 'Portal data',
  // Who may use the portal without an assigned role:
  //   'operator' every Saxon account (Enterprise application: Assignment required = No).
  //              SharePoint permissions decide who can actually change anything.
  //   'viewer'   every Saxon account, viewing only, unless assigned Operator
  //   null       only people assigned Viewer or Operator in the Enterprise application
  defaultRole: 'operator'
};
