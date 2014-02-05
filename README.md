node-xlsx-stream
================

Creates SpreadsheetML (.xlsx) files in sequence with streaming interface.

* Installation

        npm install xlsx-stream

* Features

        Multiple sheets, String, Number, Date, Duration.

* Usage

        # coffee-script
        xlsx = require "xlsx-stream"
        fs = require "fs"
        
        x = xlsx()
        x.pipe fs.createWriteStream("./out.xlsx")
        
        x.write ["foo", "bar", "buz"]
        x.write [1,2,3]
        x.write ["Date", new Date]
        x.write ["Duration", { v: 1.5, t: 'n', s:3 }]
        x.write ["Formula", {v: "ok", f: "CONCATENATE(A1,B2)"}]

        x.end()

* Multiple sheets support
        
        # coffee-script
        
        x = xlsx()
        x.pipe fs.createWriteStream("./out.xlsx")

        sheet1 = x.sheet('first sheet')
        sheet1.write ["first", "sheet"]
        sheet1.end()

        sheet2 = x.sheet('another')
        sheet2.write ["second", "sheet"]
        sheet2.end()

        x.finalize()

* Help Wanted

        Comments and Custom Number Formats
