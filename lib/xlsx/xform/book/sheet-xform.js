const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');
const {MAX_SHEET_ID} = require('../../../utils/range-guard');

class WorksheetXform extends BaseXform {
  render(xmlStream, model) {
    xmlStream.leafNode('sheet', {
      sheetId: model.id,
      name: model.name,
      state: model.state,
      'r:id': model.rId,
    });
  }

  parseOpen(node) {
    if (node.name === 'sheet') {
      const id = parseInt(node.attributes.sheetId, 10);
      if (id > MAX_SHEET_ID) {
        throw new Error(`Invalid sheetId ${id}: worksheet ids above ${MAX_SHEET_ID} are not supported.`);
      }
      this.model = {
        name: utils.xmlDecode(node.attributes.name),
        id,
        state: node.attributes.state,
        rId: node.attributes['r:id'],
      };
      return true;
    }
    return false;
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

module.exports = WorksheetXform;
