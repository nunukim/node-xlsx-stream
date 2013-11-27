module.exports =
  # node-xlsx-writer からパクってきました。
  colChar: (input)->
    input = input.toString(26)
    while input.length
      a = input.charCodeAt(input.length - 1)
      colIndex = String.fromCharCode(a + if a >= 48 and a <= 57 then 17 else -22) + colIndex
      input = if input.length > 1 then (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) else ""
    return colIndex

  escapeXML: (str)->
    String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
