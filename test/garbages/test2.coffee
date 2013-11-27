xlsx_stream = require './index'
fs = require "fs"
_ = require "lodash"
async = require "async"
from = require "from"

n = 1024*153

xlsx = xlsx_stream(level: 1)

input = from (cnt, next)->
  if cnt < n
    @emit 'data', [cnt, _.random(100)]
    process.nextTick next
  else
    @emit 'end'
  return

input.pipe(xlsx).pipe fs.createWriteStream("./foo.xlsx")

#async.eachSeries [0...n], itr, -> xlsx.end()

