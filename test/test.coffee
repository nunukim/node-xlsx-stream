xlsx_stream = require "../"
vows = require "vows"
assert = require "assert"
office = require "office"

fs = require "fs"
path = require "path"

tmp = (filename)-> path.resolve(__dirname, '../tmp', filename)

vows.describe('xlsx-stream').addBatch(
  "archiver":
    topic: ->
      zip = require("archiver").create('zip')
      zip.pipe require('concat-stream')(@callback)

      stream = require('through')()

      process.nextTick ->
        zip.append stream, name: "0.txt"
        zip.append "aaa", name: "a.txt"
        zip.append "bbb", name: "b.txt"
        zip.finalize()

        process.nextTick ->
          stream.write("ccc")
          stream.end("ddd")
      return
    "bbb": (d)->
      fs.writeFileSync("./tmp/test.zip", d)

).addBatch(
  "Array input":
    "create":
      topic: ->
        x = xlsx_stream()
        x.on 'end', @callback
        x.on 'finalize', -> console.log "FINALIZE:", arguments
        x.pipe fs.createWriteStream(tmp('array.xlsx'))
        x.write ["String", "てすと", "&'\";<>", "&amp;"]
        x.write ["Integer", 1,2,-3]
        x.write ["Float", 1.5, 0.3, 0.123456789e+23]
        x.write ["Boolean", true, false]
        x.write ["Date", new Date]
        x.write ["2 Decimals Built-in format #2", { v: 1.5, nf: '0.00' }]
        x.write ["Time Built-in format #18", { v: 1.5, nf: 'h:mm AM/PM' }]
        x.write ["Percentage Built-in format #9", { v: 0.5, nf: '0.00%' }]
        x.write ["Percentage Custom format", { v: 0.5, nf: '00.000%' }]
        x.write ["Duration 36 hours format #46", {v: 1.5, t: 'n', nf: '[h]:mm:ss' }]
        x.write ["Formula", {v: "ok", f: "CONCATENATE(A1,B2)"}]
        x.end()
        return
      "Parse xlsx":
        #topic: ->
          #office.parse(tmp('b.zip'), @callback)
          #return

        "log": (rows)->
          console.log arguments # Seems like node-xlsx does not support Inline String.
          #assert.ok()

  "Multiple sheets":
    "create":
      topic: ->
        x = xlsx_stream()
        x.on 'end', @callback
        x.on 'finalize', -> console.log "FINALIZE:", arguments
        x.pipe fs.createWriteStream(tmp('multi.xlsx'))

        sheet1 = x.sheet("1st sheet")
        sheet1.write ["This", "is", "my", "first", "worksheet"]
        sheet1.end()

        sheet2 = x.sheet("２枚目のシート")
        sheet2.write ["これが", "２枚目の", "ワークシート", "です"]
        sheet2.end()

        x.finalize()
        return

      "Parse xlsx":
        #topic: ->
          #office.parse(tmp('b.zip'), @callback)
          #return

        "log": (rows)->
          console.log arguments # Seems like node-xlsx does not support Inline String.
          # assert.ok()

).export(module, error: false)
