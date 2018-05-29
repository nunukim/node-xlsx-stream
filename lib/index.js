// Convert stream of Array to a xlsx file.

// * Usage

// out = fs.createWriteStream('out.xlsx')
// stream = xlsxStream()
// stream.pipe out

// stream.write(['aaa', 'bbb', 'ccc'])
// stream.write([1, 2, 3])
// stream.write([new Date, '090-1234-5678', 'これはテストです'])

// stream.end()
var _, archiver, duplex, sheetStream, templates, through, utils, xlsxStream;

through = require('through');

archiver = require('archiver');

_ = require('lodash');

duplex = require('duplexer');

templates = require('./templates');

utils = require("./utils");

sheetStream = require("./sheet");

module.exports = xlsxStream = function(opts = {}) {
  var defaultRepeater, defaultSheet, i, index, item, len, proxy, ref, sheets, styles, zip;
  // archiving into a zip file using archiver (internally using node's zlib built-in module)
  zip = archiver.create('zip', opts);
  defaultRepeater = through();
  proxy = duplex(defaultRepeater, zip);
  // prevent loosing data before listening 'data' event in node v0.8
  zip.pause();
  process.nextTick(function() {
    return zip.resume();
  });
  defaultSheet = null;
  sheets = [];
  styles = {
    numFmts: [
      {
        numFmtId: "0",
        formatCode: ""
      }
    ],
    cellStyleXfs: [
      {
        numFmtId: "0",
        formatCode: ""
      },
      {
        numFmtId: "1",
        formatCode: "0"
      },
      {
        numFmtId: "14",
        formatCode: "m/d/yy"
      }
    ],
    customFormatsCount: 0,
    formatCodesToStyleIndex: {}
  };
  index = 0;
  ref = styles.cellStyleXfs;
  for (i = 0, len = ref.length; i < len; i++) {
    item = ref[i];
    styles.formatCodesToStyleIndex[item.formatCode || ""] = index;
    index++;
  }
  // writing data without sheet() results in creating a default worksheet named 'Sheet1'
  defaultRepeater.once('data', function(data) {
    defaultSheet = proxy.sheet('Sheet1');
    defaultSheet.write(data);
    defaultRepeater.pipe(defaultSheet);
    return defaultRepeater.on('end', proxy.finalize);
  });
  // Append a new worksheet to the workbook
  proxy.sheet = function(name) {
    var sheet;
    index = sheets.length + 1;
    sheet = {
      index: index,
      name: name || `Sheet${index}`,
      rel: `worksheets/sheet${index}.xml`,
      path: `xl/worksheets/sheet${index}.xml`,
      styles: styles
    };
    sheets.push(sheet);
    return sheetStream(zip, sheet, opts);
  };
  // finalize the xlsx file
  proxy.finalize = function() {
    var buffer, func, j, len1, name, obj, ref1, ref2, ref3, sheet;
    // styles
    zip.append(templates.styles(styles), {
      name: "xl/styles.xml",
      store: opts.store
    });
    ref1 = templates.statics;
    // static files
    for (name in ref1) {
      buffer = ref1[name];
      zip.append(buffer, {
        name,
        store: opts.store
      });
    }
    ref2 = templates.semiStatics;
    for (name in ref2) {
      func = ref2[name];
      zip.append(func(opts), {
        name,
        store: opts.store
      });
    }
    ref3 = templates.sheet_related;
    // files modified by number of sheets
    for (name in ref3) {
      obj = ref3[name];
      buffer = obj.header;
      for (j = 0, len1 = sheets.length; j < len1; j++) {
        sheet = sheets[j];
        buffer += obj.sheet(sheet);
      }
      buffer += obj.footer;
      zip.append(buffer, {
        name,
        store: opts.store
      });
    }
    return zip.finalize(function(e, bytes) {
      if (e != null) {
        return proxy.emit('error', e);
      }
      return proxy.emit('finalize', bytes);
    });
  };
  return proxy;
};
