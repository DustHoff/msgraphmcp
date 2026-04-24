import { encodeId, escapeHtml, needsEventualConsistency, odataQuote, userPath } from '../../src/tools/shared';

describe('shared helpers', () => {
  describe('userPath', () => {
    it('returns /me for the literal "me" sentinel', () => {
      expect(userPath('me')).toBe('/me');
    });

    it('encodes UPNs for safe inclusion', () => {
      expect(userPath('alice@contoso.com')).toBe('/users/alice%40contoso.com');
    });

    it('encodes characters that could break out of the path segment', () => {
      expect(userPath('id/with/slash')).toBe('/users/id%2Fwith%2Fslash');
      expect(userPath('id?query=1')).toBe('/users/id%3Fquery%3D1');
      expect(userPath('id#fragment')).toBe('/users/id%23fragment');
    });
  });

  describe('odataQuote', () => {
    it('doubles embedded single quotes (OData escape)', () => {
      expect(odataQuote("it's")).toBe("it''s");
    });

    it('leaves other characters untouched', () => {
      expect(odataQuote('no quotes here')).toBe('no quotes here');
    });
  });

  describe('needsEventualConsistency', () => {
    it('is true when $search is set', () => {
      expect(needsEventualConsistency({ $search: '"foo"' })).toBe(true);
    });

    it('is true when $count=true', () => {
      expect(needsEventualConsistency({ $count: true })).toBe(true);
    });

    it('is false otherwise', () => {
      expect(needsEventualConsistency({ $top: 10 })).toBe(false);
    });
  });

  describe('encodeId', () => {
    it('is a no-op for GUID-like ids', () => {
      const id = '0177548a-548a-0177-8a54-77018a547701';
      expect(encodeId(id)).toBe(id);
    });

    it('encodes characters that would break out of the URL path segment', () => {
      expect(encodeId('abc/def')).toBe('abc%2Fdef');
      expect(encodeId('a?b=c')).toBe('a%3Fb%3Dc');
      expect(encodeId('a#b')).toBe('a%23b');
      expect(encodeId('a b')).toBe('a%20b');
    });
  });

  describe('escapeHtml', () => {
    it('escapes all five HTML-significant characters', () => {
      expect(escapeHtml('<script>alert("xss")&\'1\'</script>'))
        .toBe('&lt;script&gt;alert(&quot;xss&quot;)&amp;&#39;1&#39;&lt;/script&gt;');
    });

    it('escapes & so that numeric entities cannot re-introduce a tag', () => {
      // Without escaping &, "&#60;script&#62;" would be decoded by the browser
      // to "<script>" and execute — even though there is no literal `<`.
      expect(escapeHtml('&#60;script&#62;alert(1)&#60;/script&#62;'))
        .toBe('&amp;#60;script&amp;#62;alert(1)&amp;#60;/script&amp;#62;');
    });

    it('is idempotent for plain text', () => {
      expect(escapeHtml('hello world')).toBe('hello world');
    });
  });
});
