package bundles.xlsx.externs;

@:structInit
extern class Worksheet {
    public var name:String;
    @:optional public var data:Array<Array<Dynamic>>;
}