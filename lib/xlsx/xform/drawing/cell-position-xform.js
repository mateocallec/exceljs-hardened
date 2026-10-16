const BaseXform = require('../base-xform');
const IntegerXform = require('../simple/integer-xform');
const {MAX_ROWS, MAX_COLS} = require('../../../utils/range-guard');

class CellPositionXform extends BaseXform {
  constructor(options) {
    super();

    this.tag = options.tag;
    this.map = {
      'xdr:col': new IntegerXform({tag: 'xdr:col', zero: true}),
      'xdr:colOff': new IntegerXform({tag: 'xdr:colOff', zero: true}),
      'xdr:row': new IntegerXform({tag: 'xdr:row', zero: true}),
      'xdr:rowOff': new IntegerXform({tag: 'xdr:rowOff', zero: true}),
    };
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);

    this.map['xdr:col'].render(xmlStream, model.nativeCol);
    this.map['xdr:colOff'].render(xmlStream, model.nativeColOff);

    this.map['xdr:row'].render(xmlStream, model.nativeRow);
    this.map['xdr:rowOff'].render(xmlStream, model.nativeRowOff);

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model = {
          nativeCol: this.map['xdr:col'].model,
          nativeColOff: this.map['xdr:colOff'].model,
          nativeRow: this.map['xdr:row'].model,
          nativeRowOff: this.map['xdr:rowOff'].model,
        };
        // [exceljs-hardened] VULN-19: anchors are 0-based and a bottom/right anchor may sit on the
        // far edge, hence "<=". A forged <xdr:col>300000000</xdr:col> made reading
        // getImages()[0].range.tl.col allocate hundreds of MB (it walks the column list).
        if (this.model.nativeCol > MAX_COLS || this.model.nativeRow > MAX_ROWS) {
          throw new Error(
            `Invalid drawing anchor (col ${this.model.nativeCol}, row ${this.model.nativeRow}): ` +
              `Excel supports at most ${MAX_COLS} columns and ${MAX_ROWS} rows.`
          );
        }
        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }
}

module.exports = CellPositionXform;
