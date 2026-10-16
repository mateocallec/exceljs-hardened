const utils = require('../../../utils/utils');

class CacheField {
  constructor({name, sharedItems}) {
    // string type
    //
    // {
    //   'name': 'A',
    //   'sharedItems': ['a1', 'a2', 'a3']
    // }
    //
    // or
    //
    // integer type
    //
    // {
    //   'name': 'D',
    //   'sharedItems': null
    // }
    this.name = name;
    this.sharedItems = sharedItems;
  }

  render() {
    // PivotCache Field: http://www.datypic.com/sc/ooxml/e-ssml_cacheField-1.html
    // Shared Items: http://www.datypic.com/sc/ooxml/e-ssml_sharedItems-1.html

    // [exceljs-hardened] VULN-11: header names and cell values come from the worksheet (possibly
    // untrusted); they used to be interpolated raw, producing invalid XML (`R&D <x>`) or injected
    // elements (`a"/><injected foo="1`).
    const name = utils.xmlEncode(String(this.name));

    // integer types
    if (this.sharedItems === null) {
      // TK(2023-07-18): left out attributes... minValue="5" maxValue="45"
      return `<cacheField name="${name}" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" />
    </cacheField>`;
    }

    // string types
    return `<cacheField name="${name}" numFmtId="0">
      <sharedItems count="${this.sharedItems.length}">
        ${this.sharedItems.map(item => `<s v="${utils.xmlEncode(String(item))}" />`).join('')}
      </sharedItems>
    </cacheField>`;
  }
}

module.exports = CacheField;
