'use strict';

const crypto = require('crypto');

/**
 * Generate an RFC 4122 version 4 (random) UUID.
 * Replaces the 'uuid' package so the browser bundle needs no ES2021 syntax support.
 * @returns {string} A lowercase UUID, e.g. '0b8b3b7e-5f2a-4c1d-9e3f-6a7b8c9d0e1f'
 */
function uuidv4() {
  const bytes = crypto.randomBytes(16);
  // Set version (4) and variant (10xx) bits per RFC 4122 section 4.4
  /* eslint-disable no-bitwise */
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  /* eslint-enable no-bitwise */
  const hex = bytes.toString('hex');
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

module.exports = {uuidv4};
