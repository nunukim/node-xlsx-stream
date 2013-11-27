module.exports = {
  colChar: function(input) {
    var a, colIndex;
    input = input.toString(26);
    while (input.length) {
      a = input.charCodeAt(input.length - 1);
      colIndex = String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colIndex;
      input = input.length > 1 ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) : "";
    }
    return colIndex;
  },
  escapeXML: function(str) {
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }
};
