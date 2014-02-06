var archiver, duplex, sheetStream, templates, through, utils, xlsxStream, _;

through = require('through');

archiver = require('archiver');

_ = require('lodash');

duplex = require('duplexer');

templates = require('./templates');

utils = require("./utils");

sheetStream = require("./sheet");

module.exports = xlsxStream = function(opts) {
  var defaultRepeater, defaultSheet, index, item, proxy, sheets, styles, zip, _i, _len, _ref;
  if (opts == null) {
    opts = {};
  }
  zip = archiver.create('zip', opts);
  defaultRepeater = through();
  proxy = duplex(defaultRepeater, zip);
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
      }, {
        numFmtId: "1",
        formatCode: "0"
      }, {
        numFmtId: "14",
        formatCode: "m/d/yy"
      }
    ],
    customFormatsCount: 0,
    formatCodesToStyleIndex: {}
  };
  index = 0;
  _ref = styles.cellStyleXfs;
  for (_i = 0, _len = _ref.length; _i < _len; _i++) {
    item = _ref[_i];
    styles.formatCodesToStyleIndex[item.formatCode || ""] = index;
    index++;
  }
  defaultRepeater.once('data', function(data) {
    defaultSheet = proxy.sheet('Sheet1');
    defaultSheet.write(data);
    defaultRepeater.pipe(defaultSheet);
    return defaultRepeater.on('end', proxy.finalize);
  });
  proxy.sheet = function(name) {
    var sheet;
    index = sheets.length + 1;
    sheet = {
      index: index,
      name: name || ("Sheet" + index),
      rel: "worksheets/sheet" + index + ".xml",
      path: "xl/worksheets/sheet" + index + ".xml",
      styles: styles
    };
    sheets.push(sheet);
    return sheetStream(zip, sheet, opts);
  };
  proxy.finalize = function() {
    var buffer, func, name, obj, sheet, _j, _len1, _ref1, _ref2, _ref3;
    zip.append(templates.styles(styles), {
      name: "xl/styles.xml",
      store: opts.store
    });
    _ref1 = templates.statics;
    for (name in _ref1) {
      buffer = _ref1[name];
      zip.append(buffer, {
        name: name,
        store: opts.store
      });
    }
    _ref2 = templates.semiStatics;
    for (name in _ref2) {
      func = _ref2[name];
      zip.append(func(opts), {
        name: name,
        store: opts.store
      });
    }
    _ref3 = templates.sheet_related;
    for (name in _ref3) {
      obj = _ref3[name];
      buffer = obj.header;
      for (_j = 0, _len1 = sheets.length; _j < _len1; _j++) {
        sheet = sheets[_j];
        buffer += obj.sheet(sheet);
      }
      buffer += obj.footer;
      zip.append(buffer, {
        name: name,
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
