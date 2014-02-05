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

  buildCell: (ref, val)->
    return '' unless val?
    if typeof val == 'object' and !_.isDate(val)
      v = val.v
      t = val.t
      s = val.s
      f = val.f
    else
      v = val

    if _.isNumber(v) and _.isFinite(v)
      v = '<v>' + v + '</v>'
    else if _.isDate(v)
      t = 'd'
      s = '2' unless s?
      v = '<v>' + v.toISOString() + '</v>'
    else if _.isBoolean(v)
      t = 'b'
      v = '<v>' + (v ? '1' : '0') + '</v>'
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
