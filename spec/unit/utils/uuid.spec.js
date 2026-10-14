const {uuidv4} = verquire('utils/uuid');

describe('uuid', () => {
  describe('uuidv4', () => {
    it('returns an RFC 4122 version 4 UUID', () => {
      for (let i = 0; i < 100; i++) {
        expect(uuidv4()).to.match(
          /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/
        );
      }
    });

    it('returns unique values', () => {
      const ids = new Set();
      for (let i = 0; i < 1000; i++) {
        ids.add(uuidv4());
      }
      expect(ids.size).to.equal(1000);
    });
  });
});
