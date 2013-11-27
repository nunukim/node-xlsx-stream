xlsx_stream = require './index'
fs = require "fs"

out = fs.createWriteStream("./foo.xlsx")
xlsx = xlsx_stream()
xlsx.on 'data', console.log
xlsx.pipe out

xlsx.write ["あああ"]
xlsx.write ["080-1234-5678", "これはテストです"]
xlsx.write [123, 1.24, 1e50, "←数値型"]
xlsx.write ["&;\"'", new Date]

xlsx.end()

