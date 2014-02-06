_ = require "lodash"
module.exports =
  colChar: (input)->
    input = input.toString(26)
    colIndex = ''
    while input.length
      a = input.charCodeAt(input.length - 1)
      colIndex = String.fromCharCode(a + if a >= 48 and a <= 57 then 17 else -22) + colIndex
      input = if input.length > 1 then (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) else ""
    return colIndex

  escapeXML: escapeXML = (str)->
    String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
  compress: compress = (str)->
    String(str).replace(/\n\s*/g, '')

  buildCell: (ref, val, styles)->

    getStyle = (nf)->
      return unless nf
      r = styles.formatCodesToStyleIndex[nf]
      return r if r

      getBuiltinNumFmtId = (nf)->
        # ECMA-376 18.8.30
        builtin_nfs =
          'General': 0,
          '': 0,
          '0': 1,
          '0.00': 2,
          '#,##0': 3,
          '#,##0.00': 4,
          '0%': 9,
          '0.00%': 10,
          '0.00E+00': 11,
          '# ?/?': 12,
          '# ??/??': 13,
          'm/d/yy': 14, # also 30
          'd-mmm-yy': 15,
          'd-mmm': 16,
          'mmm-yy': 17,
          'h:mm AM/PM': 18,
          'h:mm:ss AM/PM': 19,
          'h:mm': 20,
          'h:mm:ss': 21,
          'm/d/yy h:mm': 22,
          '[$-404]e/m/d': 27, # also 36, 50, 57
          '#,##0 ;(#,##0)': 37,
          '#,##0 ;[Red](#,##0)': 38,
          '#,##0.00;(#,##0.00)': 39,
          '#,##0.00;[Red](#,##0.00)': 40,
          '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)': 44,
          'mm:ss': 45,
          '[h]:mm:ss': 46,
          'mmss.0': 47,
          '##0.0E+0': 48,
          '@': 49,
          't0': 59,
          't0.00': 60,
          't#,##0': 61,
          't#,##0.00': 62,
          't0%': 67,
          't0.00%': 68,
          't# ?/?': 69,
          't# ??/??': 70

        r = builtin_nfs[nf]
        return r

      numFmtId = getBuiltinNumFmtId(nf)
      unless numFmtId
        styles.customFormatsCount++
        numFmtId = 164 + styles.customFormatsCount
        styles.numFmts.push(
          numFmtId: numFmtId,
          formatCode: nf)

      s = styles.cellStyleXfs.length
      styles.cellStyleXfs.push({ numFmtId: numFmtId, formatCode: nf })
      styles.formatCodesToStyleIndex[nf] = s
      return s

    return '' unless val?
    if typeof val == 'object' and !_.isDate(val)
      v = val.v
      t = val.t
      s = val.s
      f = val.f
      s = getStyle(val.nf) if not s and val.nf
    else
      v = val

    if _.isNumber(v) and _.isFinite(v)
      v = '<v>' + v + '</v>'
      t = 'n' if val.nf and not t
    else if _.isDate(v)
      t = 'd'
      s = '2' unless s?
      v = '<v>' + v.toISOString() + '</v>'
    else if _.isBoolean(v)
      t = 'b'
      v = '<v>' + (if v is true then '1' else '0') + '</v>'
    else if v
      v = '<is><t>' + escapeXML(v) + '</t></is>'
      t = 'inlineStr'

    return '' unless v or f
    r = '<c r="' + ref + '"'
    r += ' t="' + t + '"' if t
    r += ' s="' + s + '"' if s
    r += '>'
    r += '<f>' + f + '</f>' if f
    r += v if v
    r += '</c>'
    return r
