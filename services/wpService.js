/**
 * WordPress Service
 * Handles interaction with WordPress REST API
 */
const axios = require('axios');
const { createTtlCache } = require('./asyncUtils');
const { compareEs } = require('./formatters');

// Ensure required environment variables are set
if (!process.env.WP_URL || !process.env.WP_USER || !process.env.WP_APP_PASSWORD) {
  console.error('Missing WordPress configuration in .env');
}

const wpAuth = Buffer.from(`${process.env.WP_USER}:${process.env.WP_APP_PASSWORD}`).toString('base64');
const wpInstance = axios.create({
  baseURL: `${process.env.WP_URL}/wp-json/wp/v2`,
  timeout: 5000,
  headers: {
    Authorization: `Basic ${wpAuth}`,
    'User-Agent': 'M365-App-Integration',
  },
});

const categoryCache = createTtlCache({ ttlMs: 10 * 60 * 1000, maxSize: 1 });

async function getCategories() {
  try {
    const categories = await categoryCache.get('categories', async () => {
      const response = await wpInstance.get('/categories', {
        params: { per_page: 100 },
      });
      return [...response.data].sort((a, b) => compareEs(a.name, b.name));
    });
    return categories || [];
  } catch (error) {
    console.error('Error fetching WP categories:', error.response?.data || error.message);
    return [];
  }
}

function resetCategoryCache() {
  categoryCache.clear();
}

module.exports = { getCategories, resetCategoryCache };
