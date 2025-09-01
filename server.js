// Working Teams-OpenPhone Status Sync Integration
// This works with the current OpenPhone API (no direct presence endpoints)
// Updates Teams status based on OpenPhone call events

const express = require('express');
const axios = require('axios');
const crypto = require('crypto');

const app = express();
app.use(express.json());

// Configuration from environment variables
const config = {
  teams: {
    tenantId: process.env.TEAMS_TENANT_ID,
    clientId: process.env.TEAMS_CLIENT_ID,
    clientSecret: process.env.TEAMS_CLIENT_SECRET,
  },
  openphone: {
    apiKey: process.env.OPENPHONE_API_KEY,
  },
  port: process.env.PORT || 3000,
  baseUrl: process.env.BASE_URL || `http://localhost:${process.env.PORT || 3000}`
};

// In-memory storage for active calls (in production, use a database)
const activeCalls = new Map();

// Teams Graph API authentication
async function getTeamsAccessToken() {
  try {
    const response = await axios.post(
      `https://login.microsoftonline.com/${config.teams.tenantId}/oauth2/v2.0/token`,
      new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: config.teams.clientId,
        client_secret: config.teams.clientSecret,
        scope: 'https://graph.microsoft.com/.default'
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
      }
    );
    return response.data.access_token;
  } catch (error) {
    console.error('Error getting Teams access token:', error.response?.data);
    throw error;
  }
}

function getTeamsUserFromOpenPhoneUser(openPhoneUserId) {
  const userMapping = {
    'US2gwvMWKA': 'Tyler@WindsorM.com',
  };
  
  return userMapping[openPhoneUserId];
}

// Set Teams presence
async function setTeamsPresence(userEmail, availability, activity) {
  const token = await getTeamsAccessToken();
  try {
    const response = await axios.post(
      `https://graph.microsoft.com/v1.0/users/${userEmail}/presence/setPresence`,
      {
        sessionId: crypto.randomUUID(),
        availability: availability,
        activity: activity
      },
      {
        headers: { 
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        }
      }
    );
    console.log(`Successfully updated Teams presence for ${userEmail}: ${availability}/${activity}`);
    return response.status === 200;
  } catch (error) {
    console.error('Error setting Teams presence:', error.response?.data);
    throw error;
  }
}

// Verify OpenPhone webhook signature
function verifyOpenPhoneSignature(payload, signature, secret) {
  if (!signature || !secret) {
    console.log('No signature verification - webhook secret not configured');
    return true; // Allow through if not configured (for testing)
  }
  
  // Parse the OpenPhone signature format
  const fields = signature.split(';');
  if (fields.length < 4) return false;
  
  const timestamp = fields[2];
  const providedDigest = fields[3];
  
  // Compute the data covered by the signature
  const signedData = timestamp + '.' + payload;
  
  // Convert the base64-encoded signing key to binary
  const signingKeyBinary = Buffer.from(secret, 'base64');
  
  // Compute the SHA256 HMAC digest
  const computedDigest = crypto
    .createHmac('sha256', signingKeyBinary)
    .update(signedData, 'utf8')
    .digest('base64');
  
  return crypto.timingSafeEqual(
    Buffer.from(providedDigest), 
    Buffer.from(computedDigest)
  );
}

// OpenPhone webhook handler
app.post('/webhook/openphone', async (req, res) => {
  try {
    // Get the raw payload for signature verification
    const rawPayload = JSON.stringify(req.body);
    const signature = req.headers['openphone-signature'];
    
    // Verify webhook signature (optional - for security)
    if (process.env.OPENPHONE_WEBHOOK_SECRET) {
      if (!verifyOpenPhoneSignature(rawPayload, signature, process.env.OPENPHONE_WEBHOOK_SECRET)) {
        console.error('Invalid OpenPhone webhook signature');
        return res.status(401).send('Invalid signature');
      }
    }

    const event = req.body;
    console.log(`Received OpenPhone event: ${event.type}`, {
      id: event.id,
      type: event.type,
      userId: event.data?.object?.userId
    });

    // Handle call events
    if (event.type === 'call.ringing') {
      const callData = event.data.object;
      const openPhoneUserId = callData.userId;
      const callId = callData.id;
      
      // Store active call
      activeCalls.set(callId, {
        userId: openPhoneUserId,
        startTime: new Date(),
        direction: callData.direction
      });
      
      // Get corresponding Teams user
      const teamsUser = getTeamsUserFromOpenPhoneUser(openPhoneUserId);
      
      if (teamsUser) {
        // Set Teams status to Busy when call starts ringing
        await setTeamsPresence(teamsUser, 'Busy', 'InACall');
        console.log(`Updated Teams status for ${teamsUser}: call ${callId} ringing`);
      } else {
        console.log(`No Teams user mapping found for OpenPhone user: ${openPhoneUserId}`);
      }
      
          } else if (event.type === 'call.completed') {
        const callData = event.data.object;
        const callId = callData.id;
        const openPhoneUserId = callData.userId;
        
        // Remove from active calls
        const activeCall = activeCalls.get(callId);
        activeCalls.delete(callId);
        
        // Get corresponding Teams user
        const teamsUser = getTeamsUserFromOpenPhoneUser(openPhoneUserId);
        
        if (teamsUser) {
          // Always try to set status back to Available when call completes
          // This handles cases where we missed the call.ringing event
          await setTeamsPresence(teamsUser, 'Available', 'Available');
          console.log(`Updated Teams status for ${teamsUser}: call ${callId} completed, back to available`);
        } else {
          console.log(`No Teams user mapping found for OpenPhone user: ${openPhoneUserId}`);
        }
      }
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('Error handling OpenPhone webhook:', error);
    res.status(500).send('Internal server error');
  }
});

// Teams presence change webhook handler (for future bidirectional sync)
app.post('/webhook/teams', async (req, res) => {
  try {
    // This is for future use when we want to sync Teams status changes back to OpenPhone
    // Currently OpenPhone doesn't have presence update endpoints
    console.log('Teams webhook received (not implemented yet):', req.body);
    res.status(200).send('OK');
  } catch (error) {
    console.error('Error handling Teams webhook:', error);
    res.status(500).send('Internal server error');
  }
});

// Setup OpenPhone webhooks
async function setupOpenPhoneWebhooks() {
  try {
    console.log('Setting up OpenPhone webhooks...');
    console.log('Config check:', {
      baseUrl: config.baseUrl,
      apiKey: config.openphone.apiKey ? 'Set' : 'Missing'
    });
    
    // Create webhook for call events
    const webhookData = {
      url: `${config.baseUrl}/webhook/openphone`,
      events: ['call.ringing', 'call.completed']
    };
    
    const response = await axios.post(
      'https://api.openphone.com/v1/webhooks',
      webhookData,
      {
        headers: {
          'Authorization': config.openphone.apiKey,
          'Content-Type': 'application/json'
        }
      }
    );
    
    console.log('OpenPhone webhook created successfully:', response.data.id);
    return response.data;
  } catch (error) {
    console.error('Error creating OpenPhone webhook:', error.response?.data || error.message);
    throw error;
  }
}

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ 
    status: 'ok', 
    timestamp: new Date().toISOString(),
    activeCalls: activeCalls.size
  });
});

// Get current status
app.get('/status', (req, res) => {
  res.json({
    activeCalls: Array.from(activeCalls.entries()).map(([callId, call]) => ({
      callId,
      userId: call.userId,
      startTime: call.startTime,
      direction: call.direction,
      duration: Date.now() - call.startTime.getTime()
    }))
  });
});

// Manual webhook setup endpoint (call this once after deployment) - BOTH GET AND POST
app.get('/setup-webhooks', async (req, res) => {
  try {
    const webhook = await setupOpenPhoneWebhooks();
    res.json({
      success: true,
      webhook: webhook,
      message: 'OpenPhone webhook set up successfully!'
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
      details: error.response?.data
    });
  }
});

app.post('/setup-webhooks', async (req, res) => {
  try {
    const webhook = await setupOpenPhoneWebhooks();
    res.json({
      success: true,
      webhook: webhook,
      message: 'OpenPhone webhook set up successfully!'
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
      details: error.response?.data
    });
  }
});

// Configuration check endpoint
app.get('/config-check', (req, res) => {
  const configStatus = {
    teams: {
      tenantId: !!config.teams.tenantId,
      clientId: !!config.teams.clientId,
      clientSecret: !!config.teams.clientSecret,
    },
    openphone: {
      apiKey: !!config.openphone.apiKey,
    },
    baseUrl: config.baseUrl,
    environmentCheck: {
      NODE_ENV: process.env.NODE_ENV || 'development',
      PORT: process.env.PORT || 'not set'
    }
  };
  
  res.json(configStatus);
});

// Debug endpoint to check individual environment variables
app.get('/debug-env', (req, res) => {
  res.json({
    raw_env: {
      TEAMS_TENANT_ID: process.env.TEAMS_TENANT_ID || 'undefined',
      TEAMS_CLIENT_ID: process.env.TEAMS_CLIENT_ID || 'undefined',
      TEAMS_CLIENT_SECRET: process.env.TEAMS_CLIENT_SECRET ? 'exists' : 'missing',
      OPENPHONE_API_KEY: process.env.OPENPHONE_API_KEY ? 'exists' : 'missing',
      BASE_URL: process.env.BASE_URL || 'undefined',
      NODE_ENV: process.env.NODE_ENV || 'undefined',
      PORT: process.env.PORT || 'undefined'
    }
  });
});

// Start the server
app.listen(config.port, () => {
  console.log(`Teams-OpenPhone sync service running on port ${config.port}`);
  console.log(`Webhook URL: ${config.baseUrl}`);
  console.log('Environment variables loaded:');
  console.log(`- TEAMS_TENANT_ID: ${config.teams.tenantId ? 'Set' : 'Missing'}`);
  console.log(`- TEAMS_CLIENT_ID: ${config.teams.clientId ? 'Set' : 'Missing'}`);
  console.log(`- TEAMS_CLIENT_SECRET: ${config.teams.clientSecret ? 'Set' : 'Missing'}`);
  console.log(`- OPENPHONE_API_KEY: ${config.openphone.apiKey ? 'Set' : 'Missing'}`);
  console.log(`- BASE_URL: ${config.baseUrl}`);
  console.log('After deployment, visit /setup-webhooks to configure OpenPhone webhooks');
  console.log('Visit /config-check to verify your environment variables');
});

module.exports = app;
