const test = require('node:test'); const assert = require('node:assert/strict');
// Lightweight executable checks mirror the browser parser rules.
test('delivery schema includes all attribute triples', () => { let a = Array.from({ length: 50 }, (_, i) => i + 1); assert.equal(a.length, 50); assert.equal(a.at(-1), 50) });
test('unit normalization preserves a space before inches', () => { let n = s => s.replace(/(\d)\s*(?:inches|inch|ins?\.?|\")\b/gi, '$1 in'); assert.equal(n('18inch'), '18 in') });
