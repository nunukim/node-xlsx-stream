xlsx_stream = require "../"
vows = require "vows"
assert = require "assert"
office = require "office"

fs = require "fs"
path = require "path"

tmp = (filename)-> path.resolve(__dirname, '../tmp', filename)

vows.describe('xlsx-stream').addBatch(
  "Array input":
    "create":
      topic: ->
        x = xlsx_stream()
        x.on 'end', @callback
        x.pipe fs.createWriteStream(tmp('array.xlsx'))
        x.write ["String", "てすと", "&'\";<>", "&amp;"]
        x.write ["Integer", 1,2,-3]
        x.write ["Float", 1.5, 0.3, 0.123456789e+23]
        x.write ["Date", new Date]
        x.end()
        return
      "Parse xlsx":
        topic: ->
          office.parse(tmp('array.xlsx'), @callback)
          return

        "log": (e, rows)->
          console.log arguments # Seems like node-xlsx does not support Inline String.
          assert.ok()
        #"First line (string)":
          #topic: (rows)-> rows[0]
          #"ascii string": (line)->
            #assert.equal line[0], 1
).export(module, error: false)
