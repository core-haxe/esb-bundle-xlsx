package bundles.xlsx.externs;

import js.node.Buffer;

@:jsRequire("node-xlsx", "default")
extern class Xlsx {
    public static function parse(source:Buffer):Array<Dynamic>;
    public static function build(worksheets:Array<Dynamic>):Buffer;
}