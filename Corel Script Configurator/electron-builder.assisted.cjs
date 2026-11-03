// Embute somente a raiz pública. A chave DPAPI nunca é lida pelo builder.
const base=require('./electron-builder.config.cjs');
const loadTrust=require('./build/assisted-trust.cjs');
// A configuração principal jamais incorpora uma chave do laboratório, mesmo com env antigo.
module.exports={...base,extraMetadata:{assistedUpdates:loadTrust('br.com.4printshops.desktop','internal')}};
