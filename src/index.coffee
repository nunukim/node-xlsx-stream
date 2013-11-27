# xlsx ファイルをストリームで変換する。
#
# * 使い方
#
# out = fs.createWriteStream('out.xlsx')
# stream = xlsxStream()
# stream.pipe out
#
# stream.write(['aaa', 'bbb', 'ccc'])
# stream.write([1, 2, 3])
# stream.write([new Date, '090-1234-5678', 'これはテストです'])
#
# stream.end()

through = require 'through'
zipstream = require 'zipstream'
_ = require 'lodash'
async = require 'async'
duplex = require 'duplexer'

templates = require './templates'
utils = require "./utils"

module.exports = xlsxStream = (opts = {})->
  # 列番号の26進表記(A, B, .., Z, AA, AB, ..)
  # 一度計算したらキャッシュしておく。
  colChar = _.memoize utils.colChar

  conv = (cell, val)->
    return '' unless val?
    if _.isNumber(val)
      "<c r='#{cell}'><v>#{val}</v></c>"
    else if _.isDate(val)
      "<c r='#{cell}' t='d'><v>#{val.toISOString()}</v></c>"
    else
      "<c r='#{cell}' t='inlineStr'><is><t>#{utils.escapeXML(val)}</t></is></c>"

  # 行ごとに変換してxl/worksheets/sheet1.xml に追加
  onData = (row)->
    @rowIdx ?= 1
    buf = "<row r='#{@rowIdx}'>"
    if opts.columns?
      buf += conv("#{colChar(i)}#{@rowIdx}", row[col]) for col, i in opts.columns
    else
      buf += conv("#{colChar(i)}#{@rowIdx}", val) for val, i in row
    buf += '</row>'
    @queue new Buffer(buf)
    @rowIdx++
  # sheet1.xml のフッタ部分を追加
  onEnd = ->
    @queue templates.worksheet_footer
    @queue null

  instream = through(onData, onEnd)

  # zip にアーカイブする。
  zip = zipstream.createZip(opts)

  # 入力 はinstream へ、
  # zip は出力 へ。
  proxy = duplex(instream, zip)

  # 変換された入力データは順次zip にアーカイブする。
  zip.addFile instream, name: 'xl/worksheets/sheet1.xml', ->
    # sheet1.xml が終わったら
    # 静的なファイルを追加
    addStatic = (name, cb)->
      strm = through()
      zip.addFile(strm, {name}, cb)
      strm.write(templates.statics[name])
      strm.end()
    async.eachSeries _.keys(templates.statics), addStatic, ->
      zip.finalize (bytes)-> proxy.emit 'finalize', bytes

  # sheet1.xml のヘッダ部分を書き込んでいよいよ開始。
  instream.queue templates.worksheet_header

  return proxy
