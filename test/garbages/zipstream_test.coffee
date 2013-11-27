zipstream = require 'zipstream'
through = require 'through'
fs = require 'fs'

zip = zipstream.createZip()

zip.pipe fs.createWriteStream("./aaa.zip")

t = through()
zip.addFile t, name: "bbb.txt", ->
  zip.finalize -> console.log arguments

t.write(new Buffer 'aaa')
t.write(new Buffer 'bbb')
t.end()
