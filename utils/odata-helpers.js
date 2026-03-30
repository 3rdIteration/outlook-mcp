/**
 * OData helper functions for Microsoft Graph API
 */

/**
 * Escapes a string for use in OData string literals.
 *
 * In OData, string values are enclosed in single quotes and the only
 * required escape is doubling single quotes.  We also strip control
 * characters (U+0000–U+001F) to prevent query-syntax injection, but
 * keep punctuation that is valid inside a quoted string literal (e.g.
 * hyphens, colons, slashes, etc.).
 *
 * @param {string} str - The string to escape
 * @returns {string} - The escaped string
 */
function escapeODataString(str) {
  if (!str) return str;
  
  // Replace single quotes with double single quotes (OData escaping)
  str = str.replace(/'/g, "''");
  
  // Remove control characters that could cause OData syntax errors
  // eslint-disable-next-line no-control-regex
  str = str.replace(/[\x00-\x1F\x7F]/g, '');
  
  return str;
}

/**
 * Builds an OData filter from filter conditions
 * @param {Array<string>} conditions - Array of filter conditions
 * @returns {string} - Combined OData filter expression
 */
function buildODataFilter(conditions) {
  if (!conditions || conditions.length === 0) {
    return '';
  }
  
  return conditions.join(' and ');
}

module.exports = {
  escapeODataString,
  buildODataFilter
};
