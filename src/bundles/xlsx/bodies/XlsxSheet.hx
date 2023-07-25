package bundles.xlsx.bodies;

import haxe.io.Bytes;
import haxe.Json;
import esb.core.bodies.RawBody;

@:keep
@:keepInit
@:keepSub
@:expose
@:native("bundles.xlsx.bodies.XlsxSheet")
class XlsxSheet extends RawBody {
    private var sheet:Dynamic;
    public function new(sheet:Dynamic) {
        super();

        this.sheet = sheet;
    }

    public var name(get, set):String;
    private function get_name():String {
        if (sheet == null) {
            return null;
        }
        return sheet.name;
    }
    private function set_name(value:String):String {
        if (sheet == null) {
            sheet = {};
        }
        sheet.name = value;
        return value;
    }

    public function addColumn(name:String, data:Any) {
        var first = true;
        for (row in rows()) {
            if (first) {
                row.push(name);
                first = false;
            } else {
                row.push(data);
            }
        }
    }

    public function rows():Array<Array<Any>> {
        if (sheet == null) {
            return null;
        }
        if (sheet.data == null) {
            return null;
        }
        return sheet.data;
    }

    public function row(rowIndex:Int):Array<Any> {
        if (sheet == null) {
            return null;
        }
        if (sheet.data == null) {
            return null;
        }
        return sheet.data[rowIndex];
    }

    public function data(columnIndex:Int, rowIndex:Int, value:Any = null):Any {
        if (value == null) { // getter
            if (sheet == null) {
                return null;
            }
            if (sheet.data == null) {
                return null;
            }
            var row = sheet.data[rowIndex];
            if (row == null) {
                return null;
            }
            return row[columnIndex];
        } else { // setter
            if (sheet == null) {
                sheet = {};
            }
            if (sheet.data == null) {
                sheet.data = [];
            }

            var row = sheet.data[rowIndex];
            if (row == null) {
                for (_ in 0...rowIndex + 1) {
                    sheet.data.push([]);
                }
            }

            var row = sheet.data[rowIndex];
            if (row[columnIndex] == null) {
                for (_ in 0...columnIndex + 1) {
                    row.push(null);
                }
            }
            row[columnIndex] = value;

            value = null;
        }
        return value;
    }

    public override function fromBytes(bytes:Bytes) {
        sheet = Json.parse(bytes.toString());
    }

    public override function toBytes():Bytes {
        return Bytes.ofString(Json.stringify(sheet));
    }
}