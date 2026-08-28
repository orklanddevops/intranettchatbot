const { app } = require('@azure/functions');
const { createPublicKey, verify } = require('node:crypto');

const DEFAULT_ALLOWED_ORIGIN = 'https://orkland.sharepoint.com';
const DEFAULT_REQUIRED_SCOPE = 'access_as_user';
const CLOCK_SKEW_SECONDS = 300;
const JWKS_CACHE_MS = 60 * 60 * 1000;

const jwksCache = new Map();

class AuthenticationError extends Error {}

function splitCsv(value) {
  if (!value) {
    return [];
  }

  return value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function allowedOrigins() {
  return splitCsv(process.env.CHATBOT_ALLOWED_ORIGINS || DEFAULT_ALLOWED_ORIGIN)
    .map((origin) => origin.replace(/\/$/, ''));
}

function isAllowedOrigin(origin) {
  return Boolean(origin && allowedOrigins().includes(origin.replace(/\/$/, '')));
}

function corsHeaders(origin) {
  if (!isAllowedOrigin(origin)) {
    return {};
  }

  return {
    'Access-Control-Allow-Origin': origin.replace(/\/$/, ''),
    'Access-Control-Allow-Methods': 'POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Authorization,Content-Type',
    'Access-Control-Expose-Headers': 'X-Chatbot-Proxy-Stage,X-Chatbot-Backend-Status,X-Chatbot-Backend-Layer',
    'Access-Control-Max-Age': '86400',
    Vary: 'Origin'
  };
}

function jsonResponse(payload, status, origin, extraHeaders = {}) {
  return {
    status,
    headers: {
      'Content-Type': 'application/json',
      ...corsHeaders(origin),
      ...extraHeaders
    },
    body: JSON.stringify(payload)
  };
}

function decodeBase64Url(value) {
  return Buffer.from(value, 'base64url');
}

function decodeJsonBase64Url(value) {
  return JSON.parse(decodeBase64Url(value).toString('utf8'));
}

function tenantIdFromIssuer(issuer) {
  if (!issuer) {
    return undefined;
  }

  const parts = issuer.replace(/\/$/, '').split('/');
  if (parts.length >= 4 && parts[2].toLowerCase() === 'login.microsoftonline.com') {
    return parts[3];
  }

  return undefined;
}

function acceptedAudiences() {
  const audiences = splitCsv(process.env.CHATBOT_AUTH_AUDIENCES);
  const authClientId = process.env.AUTH_CLIENT_ID;

  if (authClientId) {
    audiences.push(authClientId, `api://${authClientId}`);
  }

  return [...new Set(audiences)].sort();
}

function acceptedIssuers(tenantId) {
  const issuers = splitCsv(process.env.CHATBOT_AUTH_ISSUERS);
  const authIssuerUri = process.env.AUTH_ISSUER_URI;

  if (authIssuerUri) {
    issuers.push(authIssuerUri);
  }

  issuers.push(`https://login.microsoftonline.com/${tenantId}/v2.0`);
  issuers.push(`https://sts.windows.net/${tenantId}/`);

  return new Set(issuers.map((issuer) => issuer.replace(/\/$/, '')));
}

async function getJwks(tenantId, forceRefresh = false) {
  const cached = jwksCache.get(tenantId);
  if (!forceRefresh && cached && cached.expiresAt > Date.now()) {
    return cached.keys;
  }

  const response = await fetch(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`);
  if (!response.ok) {
    throw new AuthenticationError('Could not load token signing keys.');
  }

  const jwks = await response.json();
  const keys = Array.isArray(jwks.keys) ? jwks.keys : [];
  jwksCache.set(tenantId, {
    keys,
    expiresAt: Date.now() + JWKS_CACHE_MS
  });

  return keys;
}

async function getSigningKey(tenantId, keyId) {
  let keys = await getJwks(tenantId);
  let key = keys.find((candidate) => candidate.kid === keyId);

  if (!key) {
    keys = await getJwks(tenantId, true);
    key = keys.find((candidate) => candidate.kid === keyId);
  }

  if (!key) {
    throw new AuthenticationError('Token signing key is not trusted.');
  }

  return createPublicKey({ key, format: 'jwk' });
}

function extractBearerToken(headers) {
  const authorization = headers.get('authorization');
  if (!authorization) {
    throw new AuthenticationError('Authentication is required.');
  }

  const [scheme, token] = authorization.split(/\s+/, 2);
  if (!token || scheme.toLowerCase() !== 'bearer') {
    throw new AuthenticationError('Bearer token is required.');
  }

  return token;
}

function validateRequiredScope(claims) {
  const requiredScope = process.env.CHATBOT_REQUIRED_SCOPE || DEFAULT_REQUIRED_SCOPE;
  if (!requiredScope) {
    return;
  }

  const tokenScopes = String(claims.scp || '').split(/\s+/).filter(Boolean);
  const tokenRoles = Array.isArray(claims.roles)
    ? claims.roles
    : claims.roles
      ? [claims.roles]
      : [];

  if (!tokenScopes.includes(requiredScope) && !tokenRoles.includes(requiredScope)) {
    throw new AuthenticationError('Token does not include the required chatbot scope.');
  }
}

function validateLifetimeClaims(claims) {
  const now = Math.floor(Date.now() / 1000);

  if (typeof claims.exp !== 'number' || now > claims.exp + CLOCK_SKEW_SECONDS) {
    throw new AuthenticationError('Token has expired.');
  }

  if (typeof claims.nbf === 'number' && now + CLOCK_SKEW_SECONDS < claims.nbf) {
    throw new AuthenticationError('Token is not valid yet.');
  }
}

async function validateBearerToken(token) {
  const parts = token.split('.');
  if (parts.length !== 3) {
    throw new AuthenticationError('Invalid bearer token.');
  }

  let header;
  let claims;
  try {
    header = decodeJsonBase64Url(parts[0]);
    claims = decodeJsonBase64Url(parts[1]);
  } catch (error) {
    throw new AuthenticationError('Invalid bearer token.');
  }

  if (header.alg !== 'RS256' || !header.kid) {
    throw new AuthenticationError('Token signing algorithm is not allowed.');
  }

  const configuredTenantId = process.env.AZURE_TENANT_ID || tenantIdFromIssuer(process.env.AUTH_ISSUER_URI);
  const tokenTenantId = claims.tid;
  const tenantId = configuredTenantId || tokenTenantId;

  if (!tenantId) {
    throw new AuthenticationError('Token tenant could not be determined.');
  }

  if (configuredTenantId && tokenTenantId && tokenTenantId !== configuredTenantId) {
    throw new AuthenticationError('Token tenant is not allowed.');
  }

  const audiences = acceptedAudiences();
  if (audiences.length === 0) {
    throw new AuthenticationError('No chatbot token audience is configured.');
  }

  const tokenAudiences = Array.isArray(claims.aud) ? claims.aud : [claims.aud];
  if (!tokenAudiences.some((audience) => audiences.includes(audience))) {
    throw new AuthenticationError('Token audience is not allowed.');
  }

  const issuer = String(claims.iss || '').replace(/\/$/, '');
  if (!acceptedIssuers(tenantId).has(issuer)) {
    throw new AuthenticationError('Token issuer is not allowed.');
  }

  validateLifetimeClaims(claims);

  const signingInput = `${parts[0]}.${parts[1]}`;
  const signature = decodeBase64Url(parts[2]);
  const signingKey = await getSigningKey(tenantId, header.kid);
  const signatureValid = verify('RSA-SHA256', Buffer.from(signingInput), signingKey, signature);

  if (!signatureValid) {
    throw new AuthenticationError('Bearer token validation failed.');
  }

  validateRequiredScope(claims);
}

async function forwardToBackend(request, bearerToken, origin, context) {
  const backendUrl = process.env.BACKEND_CONVERSATION_URL;
  if (!backendUrl) {
    return jsonResponse(
      { error: 'BACKEND_CONVERSATION_URL is not configured.' },
      500,
      origin,
      { 'X-Chatbot-Proxy-Stage': 'function-config' }
    );
  }

  const timeoutSeconds = Number.parseInt(process.env.BACKEND_REQUEST_TIMEOUT_SECONDS || '230', 10);
  const requestBody = await request.text();

  try {
    const backendResponse = await fetch(backendUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${bearerToken}`,
        'Content-Type': request.headers.get('content-type') || 'application/json',
        Accept: request.headers.get('accept') || 'application/json'
      },
      body: requestBody,
      signal: AbortSignal.timeout(timeoutSeconds * 1000)
    });

    const responseBody = await backendResponse.text();
    const backendLayer = backendResponse.headers.has('x-ms-middleware-request-id') ||
      backendResponse.headers.has('www-authenticate')
      ? 'easyauth'
      : 'application';

    if (!backendResponse.ok) {
      context.warn(
        `Chatbot backend returned ${backendResponse.status} from ${backendLayer}. ` +
        `Content-Type: ${backendResponse.headers.get('content-type') || '(missing)'}. ` +
        `Body preview: ${responseBody.slice(0, 500)}`
      );
    }

    return {
      status: backendResponse.status,
      headers: {
        'Content-Type': backendResponse.headers.get('content-type') || 'application/json',
        ...corsHeaders(origin),
        'X-Chatbot-Proxy-Stage': 'backend',
        'X-Chatbot-Backend-Status': String(backendResponse.status),
        'X-Chatbot-Backend-Layer': backendLayer
      },
      body: responseBody
    };
  } catch (error) {
    context.error('Could not forward chatbot request.', error);
    return jsonResponse(
      { error: `Could not contact chatbot backend: ${error.message}` },
      502,
      origin,
      { 'X-Chatbot-Proxy-Stage': 'backend-network' }
    );
  }
}

async function conversation(request, context) {
  const origin = request.headers.get('origin');

  if (!isAllowedOrigin(origin)) {
    context.warn(`Rejected chatbot request from origin: ${origin || '(missing)'}.`);
    return jsonResponse(
      { error: 'Origin is not allowed.' },
      403,
      origin,
      { 'X-Chatbot-Proxy-Stage': 'origin' }
    );
  }

  if (request.method === 'OPTIONS') {
    return {
      status: 204,
      headers: corsHeaders(origin)
    };
  }

  try {
    const bearerToken = extractBearerToken(request.headers);
    await validateBearerToken(bearerToken);
    return await forwardToBackend(request, bearerToken, origin, context);
  } catch (error) {
    if (error instanceof AuthenticationError) {
      context.warn(`Rejected chatbot bearer token: ${error.message}`);
      return jsonResponse(
        { error: error.message },
        401,
        origin,
        { 'X-Chatbot-Proxy-Stage': 'function-auth' }
      );
    }

    context.error('Unexpected conversation proxy error.', error);
    return jsonResponse(
      { error: 'Unexpected conversation proxy error.' },
      500,
      origin,
      { 'X-Chatbot-Proxy-Stage': 'function-error' }
    );
  }
}

app.http('conversation', {
  route: 'conversation',
  methods: ['OPTIONS', 'POST'],
  authLevel: 'anonymous',
  handler: conversation
});
