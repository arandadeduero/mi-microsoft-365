/**
 * Small async and caching helpers shared across services and routes.
 */

/**
 * Run an async function over items with at most `concurrency` workers in flight.
 * Preserves input order and stops starting new work if a call rejects.
 * @param {Array} items
 * @param {number} concurrency
 * @param {Function} fn - (item, index) => Promise<any>
 * @returns {Promise<Array>}
 */
async function mapWithConcurrency(items, concurrency, fn) {
  if (items.length === 0) return [];
  const limit = Math.max(1, Math.min(concurrency, items.length));
  const results = new Array(items.length);
  let nextIndex = 0;
  let failed = false;

  async function worker() {
    while (!failed) {
      const index = nextIndex;
      nextIndex += 1;
      if (index >= items.length) return;
      try {
        results[index] = await fn(items[index], index);
      } catch (error) {
        failed = true;
        throw error;
      }
    }
  }

  await Promise.all(Array.from({ length: limit }, worker));
  return results;
}

/**
 * Bounded in-memory TTL cache with in-flight request coalescing.
 * Concurrent callers for the same key share one loader call.
 * @param {Object} options
 * @param {number} options.ttlMs - How long successful entries stay fresh
 * @param {number} [options.errorTtlMs=0] - How long to remember a failure (returns errorValue)
 * @param {*} [options.errorValue=null] - Value returned while a failure is cached
 * @param {number} [options.maxSize=200] - Maximum entries kept in memory
 * @param {Function} [options.shouldCache] - Predicate deciding whether a resolved value is cached
 */
function createTtlCache({ ttlMs, errorTtlMs = 0, errorValue = null, maxSize = 200, shouldCache = (value) => value !== null && value !== undefined } = {}) {
  const entries = new Map();
  const inflight = new Map();

  function setEntry(key, value, ttl) {
    if (entries.has(key)) entries.delete(key);
    entries.set(key, { value, expires: Date.now() + ttl });
    if (entries.size > maxSize) {
      const oldestKey = entries.keys().next().value;
      entries.delete(oldestKey);
    }
  }

  async function get(key, loader) {
    const cached = entries.get(key);
    if (cached && Date.now() < cached.expires) {
      return cached.value;
    }

    if (inflight.has(key)) {
      return inflight.get(key);
    }

    const pending = Promise.resolve()
      .then(loader)
      .then((value) => {
        if (shouldCache(value)) {
          setEntry(key, value, ttlMs);
        } else if (errorTtlMs > 0) {
          setEntry(key, value, errorTtlMs);
        }
        return value;
      })
      .catch((error) => {
        if (errorTtlMs > 0) {
          setEntry(key, errorValue, errorTtlMs);
          return errorValue;
        }
        throw error;
      })
      .finally(() => {
        inflight.delete(key);
      });

    inflight.set(key, pending);
    return pending;
  }

  function del(key) {
    entries.delete(key);
  }

  function clear() {
    entries.clear();
    inflight.clear();
  }

  return { get, delete: del, clear };
}

module.exports = { mapWithConcurrency, createTtlCache };
