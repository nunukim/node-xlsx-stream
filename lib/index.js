var async, duplex, templates, through, utils, xlsxStream, zipstream, _;

through = require('through');

zipstream = require('zipstream');

_ = require('lodash');

async = require('async');

duplex = require('duplexer');

templates = require('./templates');

utils = require("./utils");

module.exports = xlsxStream = function(opts) {
  var colChar, conv, instream, onData, onEnd, proxy, zip;
  if (opts == null) {
    opts = {};
  }
  colChar = _.memoize(utils.colChar);
  conv = function(cell, val) {
    if (val == null) {
      return '';
    }
    if (_.isNumber(val)) {
      return "<c r='" + cell + "'><v>" + val + "</v></c>";
    } else if (_.isDate(val)) {
      return "<c r='" + cell + "' t='d'><v>" + (val.toISOString()) + "</v></c>";
    } else {
      return "<c r='" + cell + "' t='inlineStr'><is><t>" + (utils.escapeXML(val)) + "</t></is></c>";
    }
  };
  onData = function(row) {
    var buf, col, i, val, _i, _j, _len, _len1, _ref;
    if (this.rowIdx == null) {
      this.rowIdx = 1;
    }
    buf = "<row r='" + this.rowIdx + "'>";
    if (opts.columns != null) {
      _ref = opts.columns;
      for (i = _i = 0, _len = _ref.length; _i < _len; i = ++_i) {
        col = _ref[i];
        buf += conv("" + (colChar(i)) + this.rowIdx, row[col]);
      }
    } else {
      for (i = _j = 0, _len1 = row.length; _j < _len1; i = ++_j) {
        val = row[i];
        buf += conv("" + (colChar(i)) + this.rowIdx, val);
      }
    }
    buf += '</row>';
    this.queue(new Buffer(buf));
    return this.rowIdx++;
  };
  onEnd = function() {
    this.queue(templates.worksheet_footer);
    return this.queue(null);
  };
  instream = through(onData, onEnd);
  zip = zipstream.createZip(opts);
  proxy = duplex(instream, zip);
  zip.addFile(instream, {
    name: 'xl/worksheets/sheet1.xml'
  }, function() {
    var addStatic;
    addStatic = function(name, cb) {
      var strm;
      strm = through();
      zip.addFile(strm, {
        name: name
      }, cb);
      strm.write(templates.statics[name]);
      return strm.end();
    };
    return async.eachSeries(_.keys(templates.statics), addStatic, function() {
      return zip.finalize(function(bytes) {
        return proxy.emit('finalize', bytes);
      });
    });
  });
  instream.queue(templates.worksheet_header);
  return proxy;
};
