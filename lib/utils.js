var compress, escapeXML, _;

_ = require("lodash");

module.exports = {
  colChar: function(input) {
    var a, colIndex;
    input = input.toString(26);
    colIndex = '';
    while (input.length) {
      a = input.charCodeAt(input.length - 1);
      colIndex = String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colIndex;
      input = input.length > 1 ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) : "";
    }
    return colIndex;
  },
  escapeXML: escapeXML = function(str) {
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  },
  compress: compress = function(str) {
    return String(str).replace(/\n\s*/g, '');
  },
  buildCell: function(ref, val) {
    var f, r, s, t, v;
    if (val == null) {
      return '';
    }
    if (typeof val === 'object' && !_.isDate(val)) {
      v = val.v;
      t = val.t;
      s = val.s;
      f = val.f;
    } else {
      v = val;
    }
    if (_.isNumber(v) && _.isFinite(v)) {
      v = '<v>' + v + '</v>';
    } else if (_.isDate(v)) {
      t = 'd';
      if (s == null) {
        s = '2';
      }
      v = '<v>' + v.toISOString() + '</v>';
    } else if (_.isBoolean(v)) {
      t = 'b';
      v = '<v>' + (v != null ? v : {
        '1': '0'
      }) + '</v>';
    } else if (v) {
      v = '<is><t>' + escapeXML(v) + '</t></is>';
      t = 'inlineStr';
    }
    if (!(v || f)) {
      return '';
    }
    r = '<c r="' + ref + '"';
    if (t) {
      r += ' t="' + t + '"';
    }
    if (s) {
      r += ' s="' + s + '"';
    }
    r += '>';
    if (f) {
      r += '<f>' + f + '</f>';
    }
    if (v) {
      r += v;
    }
    r += '</c>';
    return r;
  }
};
