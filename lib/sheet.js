var sheetStream, template, through, utils, _;

_ = require("lodash");

through = require('through');

utils = require('./utils');

template = require('./templates');
worksheetTemplates = template.worksheet;

module.exports = sheetStream = function(zip, sheet, opts) {
  var colChar, converter, nRow, onData, onEnd;
  var links = [];
  if (opts == null) {
    opts = {};
  }
  colChar = _.memoize(utils.colChar);
  nRow = 0;
  onData = function(row) {
    var buf, col, i, val, _i, _j, _len, _len1, _ref;
    nRow++;
    buf = "<row r='" + nRow + "'>";
    if (opts.columns != null) {
      _ref = opts.columns;
      for (i = _i = 0, _len = _ref.length; _i < _len; i = ++_i) {
        col = _ref[i];
        buf += utils.buildCell("" + (colChar(i)) + nRow, row[col], sheet.styles);
      }
    } else {
      for (i = _j = 0, _len1 = row.length; _j < _len1; i = ++_j) {
        val = row[i];
        if (typeof val === 'string') {
          if (val.startsWith('http://')) {
            links.push(`${colChar(i)}${nRow}-${val}`)
          }
        }
        buf += utils.buildCell("" + (colChar(i)) + nRow, val, sheet.styles);
      }
    }
    buf += '</row>';
    return this.queue(buf);
  };
  onEnd = function() {
    var converter;
    this.queue(worksheetTemplates.footer);

    if (links) {
      let rel = template.rels;
      for (name in rel) {
        func = rel[name];
        zip.append(func(links), {
          name: name
        });
      }
      this.queue(worksheetTemplates.hyperLinkStart);
      let linkCounter = 0;
      links.forEach(link => {
        linkCounter++;
        this.queue(worksheetTemplates.hyperLink(link, linkCounter));
      });
      this.queue(worksheetTemplates.hyperLinkEnd);
    }
    this.queue(worksheetTemplates.endSheet);
    this.queue(null);
    return converter = colChar = zip = null;
  };
  converter = through(onData, onEnd);
  zip.append(converter, {
    name: sheet.path,
    store: opts.store
  });
  converter.queue(worksheetTemplates.header);
  return converter;
};
