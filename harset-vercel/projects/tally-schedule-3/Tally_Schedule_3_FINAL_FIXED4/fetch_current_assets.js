'use strict';
const { fetchCurrentAssets, generateCurrentAssetsXML } = require('./fetch_tally');
module.exports = { fetchCurrentAssets, generateCurrentAssetsXML };
if (require.main === module) fetchCurrentAssets().then(() => console.log('[INFO] Done')).catch(err => { console.error('[ERROR] Failed:', err.message); process.exitCode = 1; });
