'use strict';

// Downloads the three source workbooks straight from SharePoint/OneDrive via
// Microsoft Graph, authenticating as an Azure AD app (client credentials) —
// no person signs in, so it can run on a schedule. Each file is addressed by
// its ordinary "Share" link, which Graph can resolve through /shares.

const SOURCES = [
  { key: 'cronograma', env: 'SRC_CRONOGRAMA_URL', label: 'Cronograma Oficial Board - Turmas' },
  { key: 'ocupacao', env: 'SRC_OCUPACAO_URL', label: 'Ocupação de Turmas' },
  { key: 'matriculas', env: 'SRC_MATRICULAS_URL', label: 'Planilha Master Controle de Alunos' },
];

const REQUIRED_ENV = ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', ...SOURCES.map((s) => s.env)];

function missingEnv() {
  return REQUIRED_ENV.filter((k) => !process.env[k]);
}

async function getToken() {
  const body = new URLSearchParams({
    client_id: process.env.MS_CLIENT_ID,
    client_secret: process.env.MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const r = await fetch(`https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  const json = await r.json().catch(() => ({}));
  if (!r.ok || !json.access_token) {
    throw new Error(`Microsoft login failed (${r.status}): ${json.error_description || json.error || 'unknown error'}`);
  }
  return json.access_token;
}

// Graph's documented encoding for sharing URLs: "u!" + unpadded base64url.
function shareId(url) {
  return 'u!' + Buffer.from(url, 'utf8').toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
}

async function downloadShared(url, token, label) {
  const r = await fetch(`https://graph.microsoft.com/v1.0/shares/${shareId(url)}/driveItem/content`, {
    headers: { Authorization: `Bearer ${token}` },
    redirect: 'follow',
  });
  if (!r.ok) {
    throw new Error(`could not download "${label}" (${r.status}) — check the share link and the app's Files.Read.All permission`);
  }
  return Buffer.from(await r.arrayBuffer());
}

async function fetchSources() {
  const token = await getToken();
  const out = {};
  for (const s of SOURCES) out[s.key] = await downloadShared(process.env[s.env], token, s.label);
  return out;
}

module.exports = { fetchSources, missingEnv, REQUIRED_ENV };
