var _, sheetStream, template, through, utils;

_ = require("lodash");

through = require('through');

utils = require('./utils');

template = require('./templates').worksheet;

module.exports = sheetStream = function(zip, sheet, opts = {}) {
  var colChar, converter, nRow, onData, onEnd;
  // 列番号の26進表記(A, B, .., Z, AA, AB, ..)
  // 一度計算したらキャッシュしておく。
  colChar = _.memoize(utils.colChar);
  // 行ごとに変換してxl/worksheets/sheet1.xml に追加
  nRow = 0;
  onData = function(row) {
    var buf, col, i, j, k, len, len1, ref, val;
    nRow++;
    buf = `<row r='${nRow}'>`;
    if (opts.columns != null) {
      ref = opts.columns;
      for (i = j = 0, len = ref.length; j < len; i = ++j) {
        col = ref[i];
        buf += utils.buildCell(`${colChar(i)}${nRow}`, row[col], sheet.styles);
      }
    } else {
      for (i = k = 0, len1 = row.length; k < len1; i = ++k) {
        val = row[i];
        buf += utils.buildCell(`${colChar(i)}${nRow}`, val, sheet.styles);
      }
    }
    buf += '</row>';
    return this.queue(buf);
  };
  onEnd = function() {
    var converter;
    // フッタ部分を追加
    this.queue(template.footer);
    this.queue(null);
    return converter = colChar = zip = null;
  };
  converter = through(onData, onEnd);
  zip.append(converter, {
    name: sheet.path,
    store: opts.store
  });
  // ヘッダ部分を追加
  converter.queue(template.header);
  return converter;
};
