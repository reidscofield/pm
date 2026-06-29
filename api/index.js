// Vercel serverless entry.
//
// On Vercel there is no always-on server — the platform invokes this handler for
// every request. We reuse the SAME logic as the local app: server.js exports the
// request handler and a memoized init. ensureReady() runs once per cold start
// (loads state from Supabase); handleRequest() does the actual routing.
const app = require('../server.js');

module.exports = async (req, res) => {
  await app.ensureReady();
  return app.handleRequest(req, res);
};
