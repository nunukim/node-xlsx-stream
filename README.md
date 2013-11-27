node-xlsx-stream
================

Creates SpreadsheetML (.xlsx) files in sequence with streaming interface.

* Installation

    npm install xlsx-stream

* Usage

    # coffee-script
    xlsx = require "xlsx-stream"
    fs = require "fs"

    x = xlsx()
    x.pipe fs.createWriteStream("./out.xlsx")

    x.write ["foo", "bar", "buz"]
    x.write [1,2,3]

    x.end()
