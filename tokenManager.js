const sqlite3 = require('sqlite3').verbose();
const qs = require('querystring');
const path = require('path');

class TokenManager {
  constructor(dbPath) {
    const isRender = Boolean(process.env.RENDER);

    const defaultDbPath = isRender
      ? '/tmp/tokens.db'
      : path.join(process.cwd(), 'tokens.db');

    this.dbPath = dbPath || process.env.TOKEN_DB_PATH || defaultDbPath;

    this.clientId = process.env.BB_CLIENT_ID;
    this.clientSecret = process.env.BB_CLIENT_SECRET;
    this.bootstrapRefreshToken = process.env.BB_REFRESH_TOKEN;

    if (!this.clientId || !this.clientSecret) {
      throw new Error('Missing required environment variables: BB_CLIENT_ID and/or BB_CLIENT_SECRET.');
    }

    this.db = null;
    this.initPromise = this._initDb();

    console.log(`Token manager initialized. Storage path: ${this.dbPath}`);
  }

  async fetch(...args) {
    const { default: fetch } = await import('node-fetch');
    return fetch(...args);
  }

  // ---------------------------------------------------------
  // Initialize SQLite DB
  // ---------------------------------------------------------
  async _initDb() {
    return new Promise((resolve, reject) => {
      this.db = new sqlite3.Database(this.dbPath, (err) => {
        if (err) {
          console.error('Failed to open token database.');
          return reject(err);
        }

        this.db.run(
          `
          CREATE TABLE IF NOT EXISTS tokens (
            id INTEGER PRIMARY KEY,
            refresh_token TEXT,
            access_token TEXT,
            expires_at INTEGER
          )
          `,
          (tableErr) => {
            if (tableErr) {
              return reject(tableErr);
            }

            console.log('Token database initialized.');
            resolve();
          }
        );
      });
    });
  }

  // ---------------------------------------------------------
  // Save tokens (supports rotating refresh token flow)
  // ---------------------------------------------------------
  async saveTokens(accessToken, refreshToken, expiresIn) {
    await this.initPromise;

    const expiresAt = Math.floor(Date.now() / 1000) + expiresIn;

    return new Promise((resolve, reject) => {
      this.db.serialize(() => {
        this.db.run('DELETE FROM tokens');
        this.db.run(
          'INSERT INTO tokens (refresh_token, access_token, expires_at) VALUES (?, ?, ?)',
          [refreshToken, accessToken, expiresAt],
          (err) => {
            if (err) {
              reject(err);
            } else {
              console.log(`Tokens saved. Access token expires in approximately ${expiresIn} seconds.`);
              resolve();
            }
          }
        );
      });
    });
  }

  // ---------------------------------------------------------
  // Read tokens from DB
  // ---------------------------------------------------------
  async getTokens() {
    await this.initPromise;

    return new Promise((resolve, reject) => {
      this.db.get(
        'SELECT access_token, refresh_token, expires_at FROM tokens LIMIT 1',
        [],
        (err, row) => {
          if (err) {
            reject(err);
          } else {
            resolve(row || null);
          }
        }
      );
    });
  }

  // ---------------------------------------------------------
  // Refresh OAuth tokens using refresh_token
  // ---------------------------------------------------------
  async refreshAccessToken(refreshToken) {
    const tokenUrl = 'https://api.bluebeam.com/oauth2/token';

    const payload = {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: this.clientId,
      client_secret: this.clientSecret
    };

    console.log('Refreshing OAuth access token...');

    const response = await this.fetch(tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: qs.stringify(payload)
    });

    const text = await response.text();

    if (!response.ok) {
      console.error(`Token refresh failed with status ${response.status}.`);
      throw new Error(`Token refresh failed: ${response.status}`);
    }

    let data;
    try {
      data = JSON.parse(text);
    } catch (err) {
      throw new Error('Token refresh response could not be parsed as JSON.');
    }

    if (!data.access_token || !data.refresh_token || !data.expires_in) {
      throw new Error('Token refresh response did not include expected token fields.');
    }

    console.log('OAuth token refresh successful.');
    return data;
  }

  // ---------------------------------------------------------
  // Get valid access token (handles refresh + bootstrap)
  // ---------------------------------------------------------
  async getValidAccessToken() {
    await this.initPromise;

    const tokens = await this.getTokens();
    const nowUnix = Math.floor(Date.now() / 1000);

    // Case 1: No tokens in DB → bootstrap using env refresh token
    if (!tokens) {
      console.log('No cached tokens found. Attempting bootstrap using BB_REFRESH_TOKEN...');

      if (!this.bootstrapRefreshToken) {
        throw new Error('No stored tokens found and BB_REFRESH_TOKEN is not configured.');
      }

      const newTokens = await this.refreshAccessToken(this.bootstrapRefreshToken);

      await this.saveTokens(
        newTokens.access_token,
        newTokens.refresh_token,
        newTokens.expires_in
      );

      console.log('Initial token bootstrap completed.');
      return newTokens.access_token;
    }

    // Case 2: Access token still valid
    if (tokens.expires_at > nowUnix + 300) {
      console.log('Using cached access token.');
      return tokens.access_token;
    }

    // Case 3: Access token expired → refresh with stored refresh token
    console.log('Cached access token is near expiry or expired. Refreshing...');

    try {
      const newTokens = await this.refreshAccessToken(tokens.refresh_token);

      await this.saveTokens(
        newTokens.access_token,
        newTokens.refresh_token,
        newTokens.expires_in
      );

      console.log('Stored refresh token used successfully.');
      return newTokens.access_token;
    } catch (err) {
      console.warn('Stored refresh token failed. Attempting fallback bootstrap refresh token...');

      if (!this.bootstrapRefreshToken) {
        throw new Error('Stored refresh failed and BB_REFRESH_TOKEN is not configured.');
      }

      const newTokens = await this.refreshAccessToken(this.bootstrapRefreshToken);

      await this.saveTokens(
        newTokens.access_token,
        newTokens.refresh_token,
        newTokens.expires_in
      );

      console.log('Recovered access using fallback refresh token.');
      return newTokens.access_token;
    }
  }

  close() {
    if (this.db) {
      this.db.close();
    }
  }
}

module.exports = TokenManager;
