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
archiver = require 'archiver'
_ = require 'lodash'
async = require 'async'
duplex = require 'duplexer'

templates = require './templates'
utils = require "./utils"

module.exports = xlsxStream = (opts = {})->
  # 列番号の26進表記(A, B, .., Z, AA, AB, ..)
  # 一度計算したらキャッシュしておく。
  colChar = _.memoize utils.colChar

  # zip にアーカイブする。
  zip = archiver.create('zip', opts)

  # prevent loosing data before listening 'data' event in v0.8
  zip.pause()
  process.nextTick -> zip.resume()

  # 静的なファイル
  for name, buffer of templates.statics
    zip.append buffer, {name, store: opts.store}

  # 行ごとに変換してxl/worksheets/sheet1.xml に追加
  onData = (row)->
    unless @rowIdx?
      @queue templates.worksheet_header # sheet1.xml のヘッダ部分
      @rowIdx = 1

    buf = "<row r='#{@rowIdx}'>"
    if opts.columns?
      buf += utils.buildCell("#{colChar(i)}#{@rowIdx}", row[col]) for col, i in opts.columns
    else
      buf += utils.buildCell("#{colChar(i)}#{@rowIdx}", val) for val, i in row
    buf += '</row>'
    @queue new Buffer(buf)
    @rowIdx++

  # sheet1.xml のフッタ部分を追加
  onEnd = ->
    @queue templates.worksheet_header unless @rowIdx
    @queue templates.worksheet_footer
    @queue null

  instream = through(onData, onEnd)

  # 変換された入力データは順次zip にアーカイブする。
  zip.append instream, {name: 'xl/worksheets/sheet1.xml', store: opts.store}

  zip.finalize (e, bytes)->
    return proxy.emit 'error', e if e?
    proxy.emit 'finalize', bytes

  # 入力 はinstream へ、
  # zip は出力 へ。
  proxy = duplex(instream, zip)

  return proxy
