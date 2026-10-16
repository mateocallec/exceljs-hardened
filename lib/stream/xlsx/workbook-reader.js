const fs = require('fs');
const {EventEmitter} = require('events');
const {PassThrough, Readable} = require('readable-stream');
const nodeStream = require('stream');
const unzip = require('unzipper');
const tmp = require('tmp');
const iterateStream = require('../../utils/iterate-stream');
const parseSax = require('../../utils/parse-sax');
const {createStreamLimiter} = require('../../utils/decompression-guard');

const StyleManager = require('../../xlsx/xform/style/styles-xform');
const WorkbookXform = require('../../xlsx/xform/book/workbook-xform');
const RelationshipsXform = require('../../xlsx/xform/core/relationships-xform');

const WorksheetReader = require('./worksheet-reader');
const HyperlinkReader = require('./hyperlink-reader');

tmp.setGracefulCleanup();

class WorkbookReader extends EventEmitter {
  constructor(input, options = {}) {
    super();

    this.input = input;

    this.options = {
      worksheets: 'emit',
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'ignore',
      entries: 'ignore',
      ...options,
    };

    this.styles = new StyleManager();
    this.styles.init();
  }

  _getStream(input) {
    if (input instanceof nodeStream.Readable || input instanceof Readable) {
      return input;
    }
    if (typeof input === 'string') {
      return fs.createReadStream(input);
    }
    throw new Error(`Could not recognise input: ${input}`);
  }

  async read(input, options) {
    try {
      for await (const {eventType, value} of this.parse(input, options)) {
        switch (eventType) {
          case 'shared-strings':
            this.emit(eventType, value);
            break;
          case 'worksheet':
            this.emit(eventType, value);
            await value.read();
            break;
          case 'hyperlinks':
            this.emit(eventType, value);
            break;
        }
      }
      this.emit('end');
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }

  async *[Symbol.asyncIterator]() {
    for await (const {eventType, value} of this.parse()) {
      if (eventType === 'worksheet') {
        yield value;
      }
    }
  }

  async *parse(input, options) {
    if (options) this.options = options;
    const stream = (this.stream = this._getStream(input || this.input));
    const zip = unzip.Parse({forceStream: true});
    stream.pipe(zip);

    // [exceljs-hardened] VULN-22: whatever happens (malformed sheet, corrupt entry, consumer leaving the
    // iteration early), every temp file created below is removed and the open read stream is closed.
    // Before, the error path leaked one tmp-* file per remaining sheet (and file descriptors).
    let currentFileStream;
    // worksheets, deferred for parsing after shared strings reading
    const waitingWorkSheets = [];
    try {
      // [exceljs-hardened] VULN-07: opt-in limits on the real decompressed size
      const limiter = createStreamLimiter(this.options);
      const abort = error => {
        stream.unpipe(zip);
        zip.destroy(error); // surfaces the error in the loop below
        if (typeof stream.destroy === 'function') stream.destroy();
      };

      for await (const rawEntry of iterateStream(zip)) {
        const entry = limiter ? limiter(rawEntry, abort) : rawEntry;
        let match;
        let sheetNo;
        switch (entry.path) {
          case '_rels/.rels':
            break;
          case 'xl/_rels/workbook.xml.rels':
            await this._parseRels(entry);
            break;
          case 'xl/workbook.xml':
            await this._parseWorkbook(entry);
            break;
          case 'xl/sharedStrings.xml':
            yield* this._parseSharedStrings(entry);
            break;
          case 'xl/styles.xml':
            await this._parseStyles(entry);
            break;
          default:
            if (entry.path.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
              match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
              sheetNo = match[1];
              if (this.sharedStrings && this.workbookRels) {
                yield* this._parseWorksheet(iterateStream(entry), sheetNo);
              } else {
                // create temp file for each worksheet.
                // The file is created synchronously on purpose: the zip parser only moves on to the
                // next entry once this one is consumed, and an async gap here could make it end
                // before xl/sharedStrings.xml was reached (worksheet parsed without shared strings).
                const {name: path, removeCallback: tempFileCleanupCallback} = tmp.fileSync({
                  discardDescriptor: true,
                });
                waitingWorkSheets.push({sheetNo, path, tempFileCleanupCallback});
                await new Promise((resolve, reject) => {
                  const tempStream = fs.createWriteStream(path);
                  tempStream.on('error', reject);
                  // [exceljs-hardened] VULN-23: a corrupt deflate stream errors on the entry, not on
                  // tempStream: without this the promise (and the whole iteration) never settled.
                  entry.on('error', error => {
                    tempStream.destroy();
                    reject(error);
                  });
                  entry.pipe(tempStream);
                  tempStream.on('finish', resolve);
                });
              }
            } else if (entry.path.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
              match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
              sheetNo = match[1];
              yield* this._parseHyperlinks(iterateStream(entry), sheetNo);
            }
            break;
        }
        entry.autodrain();
      }

      for (const waiting of waitingWorkSheets) {
        const {sheetNo, path} = waiting;
        let fileStream = fs.createReadStream(path);
        currentFileStream = fileStream;
        // TODO: Remove once node v8 is deprecated
        // Detect and upgrade old fileStreams
        if (!fileStream[Symbol.asyncIterator]) {
          fileStream = fileStream.pipe(new PassThrough());
        }
        yield* this._parseWorksheet(fileStream, sheetNo);
        fileStream.destroy();
        waiting.tempFileCleanupCallback();
        waiting.cleaned = true;
      }
    } finally {
      if (currentFileStream) currentFileStream.destroy();
      for (const waiting of waitingWorkSheets) {
        if (!waiting.cleaned) {
          waiting.cleaned = true;
          try {
            waiting.tempFileCleanupCallback();
          } catch (e) {
            // already removed
          }
        }
      }
    }
  }

  _emitEntry(payload) {
    if (this.options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  async _parseRels(entry) {
    const xform = new RelationshipsXform();
    this.workbookRels = await xform.parseStream(iterateStream(entry));
  }

  async _parseWorkbook(entry) {
    this._emitEntry({type: 'workbook'});

    const workbook = new WorkbookXform();
    await workbook.parseStream(iterateStream(entry));

    this.properties = workbook.map.workbookPr;
    this.model = workbook.model;
  }

  async *_parseSharedStrings(entry) {
    this._emitEntry({type: 'shared-strings'});
    switch (this.options.sharedStrings) {
      case 'cache':
        this.sharedStrings = [];
        break;
      case 'emit':
        break;
      default:
        return;
    }

    let text = null;
    let richText = [];
    let index = 0;
    let font = null;
    for await (const events of parseSax(iterateStream(entry))) {
      for (const {eventType, value} of events) {
        if (eventType === 'opentag') {
          const node = value;
          switch (node.name) {
            case 'b':
              font = font || {};
              font.bold = true;
              break;
            case 'charset':
              font = font || {};
              font.charset = parseInt(node.attributes.charset, 10);
              break;
            case 'color':
              font = font || {};
              font.color = {};
              if (node.attributes.rgb) {
                font.color.argb = node.attributes.argb;
              }
              if (node.attributes.val) {
                font.color.argb = node.attributes.val;
              }
              if (node.attributes.theme) {
                font.color.theme = node.attributes.theme;
              }
              break;
            case 'family':
              font = font || {};
              font.family = parseInt(node.attributes.val, 10);
              break;
            case 'i':
              font = font || {};
              font.italic = true;
              break;
            case 'outline':
              font = font || {};
              font.outline = true;
              break;
            case 'rFont':
              font = font || {};
              font.name = node.value;
              break;
            case 'si':
              font = null;
              richText = [];
              text = null;
              break;
            case 'sz':
              font = font || {};
              font.size = parseInt(node.attributes.val, 10);
              break;
            case 'strike':
              break;
            case 't':
              text = null;
              break;
            case 'u':
              font = font || {};
              font.underline = true;
              break;
            case 'vertAlign':
              font = font || {};
              font.vertAlign = node.attributes.val;
              break;
          }
        } else if (eventType === 'text') {
          text = text ? text + value : value;
        } else if (eventType === 'closetag') {
          const node = value;
          switch (node.name) {
            case 'r':
              richText.push({
                font,
                text,
              });

              font = null;
              text = null;
              break;
            case 'si':
              if (this.options.sharedStrings === 'cache') {
                this.sharedStrings.push(richText.length ? {richText} : text);
              } else if (this.options.sharedStrings === 'emit') {
                yield {index: index++, text: richText.length ? {richText} : text};
              }

              richText = [];
              font = null;
              text = null;
              break;
          }
        }
      }
    }
  }

  async _parseStyles(entry) {
    this._emitEntry({type: 'styles'});
    if (this.options.styles === 'cache') {
      this.styles = new StyleManager();
      await this.styles.parseStream(iterateStream(entry));
    }
  }

  *_parseWorksheet(iterator, sheetNo) {
    this._emitEntry({type: 'worksheet', id: sheetNo});
    const worksheetReader = new WorksheetReader({
      workbook: this,
      id: sheetNo,
      iterator,
      options: this.options,
    });

    const matchingRel = (this.workbookRels || []).find(rel => rel.Target === `worksheets/sheet${sheetNo}.xml`);
    const matchingSheet = matchingRel && (this.model.sheets || []).find(sheet => sheet.rId === matchingRel.Id);
    if (matchingSheet) {
      worksheetReader.id = matchingSheet.id;
      worksheetReader.name = matchingSheet.name;
      worksheetReader.state = matchingSheet.state;
    }
    if (this.options.worksheets === 'emit') {
      yield {eventType: 'worksheet', value: worksheetReader};
    }
  }

  *_parseHyperlinks(iterator, sheetNo) {
    this._emitEntry({type: 'hyperlinks', id: sheetNo});
    const hyperlinksReader = new HyperlinkReader({
      workbook: this,
      id: sheetNo,
      iterator,
      options: this.options,
    });
    if (this.options.hyperlinks === 'emit') {
      yield {eventType: 'hyperlinks', value: hyperlinksReader};
    }
  }
}

// for reference - these are the valid values for options
WorkbookReader.Options = {
  worksheets: ['emit', 'ignore'],
  sharedStrings: ['cache', 'emit', 'ignore'],
  hyperlinks: ['cache', 'emit', 'ignore'],
  styles: ['cache', 'ignore'],
  entries: ['emit', 'ignore'],
};

module.exports = WorkbookReader;
