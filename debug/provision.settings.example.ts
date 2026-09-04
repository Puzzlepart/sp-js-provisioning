/**
 * Auth settings for the local provisioning runner (debug/provision.ts).
 *
 * Copy this file to `debug/provision.settings.ts` (gitignored) and fill in your
 * Azure AD app registration. The app needs SharePoint application permission
 * `Sites.FullControl.All` (admin consented) to provision sites.
 *
 * Use EITHER a client secret OR a certificate.
 *  - Certificate (recommended): create a self-signed cert, upload the .cer to the
 *    app registration, and point `privateKeyPath` at the PEM private key
 *    (-----BEGIN PRIVATE KEY-----). `thumbprint` is the cert's SHA-1 (hex).
 *  - Client secret: simpler to set up but secret-based app-only is blocked on
 *    many tenants — use the certificate if the secret path fails to authenticate.
 */
const settings = {
  tenantId: '<tenant-guid-or-domain.onmicrosoft.com>',
  clientId: '<app-registration-client-id>',

  // Option A — client secret
  clientSecret: '',

  // Option B — certificate (leave clientSecret empty to use this)
  certificate: {
    thumbprint: '',
    privateKeyPath: '' // e.g. /Users/you/certs/pp-provision.pem
  }
}

export default settings
