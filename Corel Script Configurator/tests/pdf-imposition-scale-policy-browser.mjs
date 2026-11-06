process.env.NUP_POLICY_ONLY='1';
process.env.NUP_SELECTION_RESULTS??='.desktop-test/nup26/'+(process.env.NUP_SELECTION_ELECTRON==='1'?'electron':'chrome');
await import('./pdf-imposition-selection-browser.mjs');
