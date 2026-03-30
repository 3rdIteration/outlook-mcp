const crypto = require('crypto');

const DEFAULT_TTL_MS = 10 * 60 * 1000;
const DEFAULT_MAX_ENTRIES = 100;

function toBase64Url(buffer) {
  return buffer
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

function generatePkcePair() {
  const codeVerifier = toBase64Url(crypto.randomBytes(32));
  const codeChallenge = toBase64Url(
    crypto.createHash('sha256').update(codeVerifier).digest()
  );

  return {
    codeVerifier,
    codeChallenge
  };
}

function createOAuthTransactionStore(options = {}) {
  const ttlMs = options.ttlMs || DEFAULT_TTL_MS;
  const maxEntries = options.maxEntries || DEFAULT_MAX_ENTRIES;
  const transactions = new Map();

  function pruneExpiredTransactions() {
    const now = Date.now();

    for (const [state, transaction] of transactions.entries()) {
      if ((now - transaction.createdAt) > ttlMs) {
        transactions.delete(state);
      }
    }

    while (transactions.size > maxEntries) {
      const oldestState = transactions.keys().next().value;
      transactions.delete(oldestState);
    }
  }

  function createTransaction(metadata = {}) {
    pruneExpiredTransactions();

    const state = crypto.randomBytes(32).toString('hex');
    const { codeVerifier, codeChallenge } = generatePkcePair();

    transactions.set(state, {
      createdAt: Date.now(),
      codeVerifier,
      metadata
    });

    return {
      state,
      codeVerifier,
      codeChallenge
    };
  }

  function consumeTransaction(state) {
    pruneExpiredTransactions();

    if (!state) {
      return null;
    }

    const transaction = transactions.get(state);
    if (!transaction) {
      return null;
    }

    transactions.delete(state);
    return transaction;
  }

  return {
    createTransaction,
    consumeTransaction
  };
}

module.exports = {
  createOAuthTransactionStore,
  generatePkcePair
};
