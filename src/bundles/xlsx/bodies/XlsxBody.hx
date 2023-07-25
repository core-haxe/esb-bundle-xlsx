package bundles.xlsx.bodies;

import bundles.xlsx.externs.Worksheet;
import js.node.Buffer;
import haxe.io.Bytes;
import esb.core.bodies.RawBody;

/*
#if !xlsx_impl

@:jsRequire("./bundle-xlsx.js", "esb.core.bodies.XlsxBody")
@:native("esb.core.bodies.XlsxBody")
extern class XlsxBody extends RawBody {
}

#else
*/

@:keep
@:keepInit
@:keepSub
@:expose
@:native("bundles.xlsx.bodies.XlsxBody")
class XlsxBody extends RawBody {
    private var worksheets:Array<Dynamic> = [];

    public function listSheets():Array<String> {
        var names = [];
        for (sheet in worksheets) {
            names.push(sheet.name);
        }
        return names;
    }

    public function sheet(name:String, autoCreate:Bool = true):XlsxSheet {
        var foundSheetData:Dynamic = null;
        for (worksheet in worksheets) {
            if (worksheet.name == name) {
                foundSheetData = worksheet;
                break;
            }
        }

        if (foundSheetData == null && autoCreate) {
            foundSheetData = { name: name, data: [] };
            worksheets.push(foundSheetData);
        }

        return new XlsxSheet(foundSheetData);
    }

    public function sheets():Array<XlsxSheet> {
        var array = [];
        for (worksheet in worksheets) {
            array.push(new XlsxSheet(worksheet));
        }
        return array;
    }

    public override function fromBytes(bytes:Bytes) {
        worksheets = bundles.xlsx.externs.Xlsx.parse(Buffer.hxFromBytes(bytes));
    }

    public override function toBytes():Bytes {
        var buffer = bundles.xlsx.externs.Xlsx.build(worksheets);
        return buffer.hxToBytes();
    }
}

//#end