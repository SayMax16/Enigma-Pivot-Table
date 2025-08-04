const enigma = require('enigma.js');
const WebSocket = require('ws');
const path = require('path');
const fs = require('fs');

class SessionManager {
  constructor(config) {
    this.config = config;
    this.session = null;
    this.global = null;
    this.doc = null;
  }

  // Helper function to read certificate files
  readCert(filename) {
    return fs.readFileSync(path.resolve(this.config.certificatesPath, filename));
  }

  // Create enigma.js session with optimized configuration
  createSession() {
    const schema = require('enigma.js/schemas/12.20.0.json');
    
    const sessionConfig = {
      schema,
      url: `wss://${this.config.engineHost}:${this.config.enginePort}/app/${this.config.appId}`,
      createSocket: (url) => {
        const socketOptions = {};

        // Configure authentication based on method
        if (this.config.authMethod === 'certificates') {
          socketOptions.ca = [this.readCert('root.pem')];
          socketOptions.key = this.readCert('client_key.pem');
          socketOptions.cert = this.readCert('client.pem');
          socketOptions.rejectUnauthorized = false; // Disable certificate hostname verification
          socketOptions.headers = {
            'X-Qlik-User': `UserDirectory=${encodeURIComponent(this.config.userDirectory)}; UserId=${encodeURIComponent(this.config.userId)}`,
          };
        } else if (this.config.authMethod === 'apikey') {
          socketOptions.headers = {
            Authorization: `Bearer ${this.config.apiKey}`,
          };
        } else if (this.config.authMethod === 'jwt') {
          socketOptions.headers = {
            Authorization: `Bearer ${this.config.jwtToken}`,
          };
        }

        return new WebSocket(url, socketOptions);
      },
      // Optimization: Add retry interceptor for aborted requests
      responseInterceptors: [{
        onRejected: function retryAbortedError(sessionReference, request, error) {
          console.log('Request rejected:', error.message);
          
          // Retry aborted requests (common during heavy calculations)
          if (error.code === schema.enums.LocalizedErrorCode.LOCERR_GENERIC_ABORTED) {
            request.tries = (request.tries || 0) + 1;
            console.log(`Retry attempt #${request.tries}`);
            
            if (request.tries <= 3) { // Max 3 retries
              return request.retry();
            }
          }
          
          return this.Promise.reject(error);
        },
      }],
      // Enable delta protocol for bandwidth optimization (default: true)
      protocol: {
        delta: true,
      },
    };

    this.session = enigma.create(sessionConfig);
    return this.session;
  }

  // Open session and connect to app
  async connect() {
    try {
      if (!this.session) {
        this.createSession();
      }

      console.log('Opening session...');
      this.global = await this.session.open();
      console.log('Session opened successfully');

      console.log('Opening document...');
      this.doc = await this.global.openDoc(this.config.appId);
      console.log('Document opened successfully');

      return {
        session: this.session,
        global: this.global,
        doc: this.doc,
      };
    } catch (error) {
      console.error('Failed to connect:', error);
      throw error;
    }
  }

  // Get document reference
  getDoc() {
    if (!this.doc) {
      throw new Error('Document not connected. Call connect() first.');
    }
    return this.doc;
  }

  // Clean shutdown
  async close() {
    try {
      if (this.session) {
        console.log('Closing session...');
        await this.session.close();
        console.log('Session closed successfully');
      }
    } catch (error) {
      console.error('Error closing session:', error);
    } finally {
      this.session = null;
      this.global = null;
      this.doc = null;
    }
  }

  // Monitor session traffic (for debugging)
  enableTrafficLogging() {
    if (this.session) {
      this.session.on('traffic:sent', (data) => {
        console.log('Sent:', JSON.stringify(data, null, 2));
      });
      this.session.on('traffic:received', (data) => {
        console.log('Received:', JSON.stringify(data, null, 2));
      });
    }
  }
}

module.exports = SessionManager;