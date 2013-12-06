var archiver, async, duplex, templates, through, utils, xlsxStream, _;

through = require('through');

archiver = require('archiver');

_ = require('lodash');

async = require('async');

duplex = require('duplexer');

templates = require('./templates');

utils = require("./utils");

module.exports = xlsxStream = function(opts) {
  var buffer, colChar, instream, name, onData, onEnd, proxy, zip, _ref;
  if (opts == null) {
    opts = {};
  }
  colChar = _.memoize(utils.colChar);
  zip = archiver.create('zip', opts);
  zip.pause();
  process.nextTick(function() {
    return zip.resume();
  });
  _ref = templates.statics;
  for (name in _ref) {
    buffer = _ref[name];
    zip.append(buffer, {
      name: name,
      store: opts.store
    });
  }
  onData = function(row) {
    var buf, col, i, val, _i, _j, _len, _len1, _ref1;
    if (this.rowIdx == null) {
      this.queue(templates.worksheet_header);
      this.rowIdx = 1;
    }
    buf = "<row r='" + this.rowIdx + "'>";
    if (opts.columns != null) {
      _ref1 = opts.columns;
      for (i = _i = 0, _len = _ref1.length; _i < _len; i = ++_i) {
        col = _ref1[i];
        buf += utils.buildCell("" + (colChar(i)) + this.rowIdx, row[col]);
      }
    } else {
      for (i = _j = 0, _len1 = row.length; _j < _len1; i = ++_j) {
        val = row[i];
        buf += utils.buildCell("" + (colChar(i)) + this.rowIdx, val);
      }
    }
    buf += '</row>';
    this.queue(new Buffer(buf));
    return this.rowIdx++;
  };
  onEnd = function() {
    if (!this.rowIdx) {
      this.queue(templates.worksheet_header);
    }
    this.queue(templates.worksheet_footer);
    return this.queue(null);
  };
  instream = through(onData, onEnd);
  zip.append(instream, {
    name: 'xl/worksheets/sheet1.xml',
    store: opts.store
  });
  zip.finalize(function(e, bytes) {
    if (e != null) {
      return proxy.emit('error', e);
    }
    return proxy.emit('finalize', bytes);
  });
  proxy = duplex(instream, zip);
  return proxy;
};
