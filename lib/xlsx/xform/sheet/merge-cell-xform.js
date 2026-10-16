const BaseXform = require('../base-xform');
const colCache = require('../../../utils/col-cache');
const {createRangeBudget, DEFAULT_MAX_MERGE_CELLS} = require('../../../utils/range-guard');

class MergeCellXform extends BaseXform {
  constructor(options) {
    super();
    // one instance per worksheet parse: the budget is cumulative per sheet
    this.budget = createRangeBudget(
      (options && options.maxMergeCells) || DEFAULT_MAX_MERGE_CELLS,
      'merged cells',
      'maxMergeCells'
    );
  }

  get tag() {
    return 'mergeCell';
  }

  render(xmlStream, model) {
    xmlStream.leafNode('mergeCell', {ref: model});
  }

  parseOpen(node) {
    if (node.name === 'mergeCell') {
      this.model = node.attributes.ref;
      this._checkBudget(this.model);
      return true;
    }
    return false;
  }

  _checkBudget(ref) {
    let decoded;
    try {
      decoded = colCache.decode(ref);
    } catch (e) {
      return; // malformed ref: keep the historical behaviour (reported later, when the Range is built)
    }
    this.budget.add(decoded);
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

module.exports = MergeCellXform;
