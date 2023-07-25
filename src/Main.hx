package;

import esb.core.Bus.*;
import esb.core.bodies.CsvBody;
import esb.core.Message;

class Main {
    static function main() {
        trace(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> XLXS BUNDLE MAIN RUNNING");
        registerMessageType(bundles.xlsx.bodies.XlsxBody, () -> {
            trace(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> CREATING XLSX");
            var m = @:privateAccess new Message<bundles.xlsx.bodies.XlsxBody>();
            m.body = new bundles.xlsx.bodies.XlsxBody();
            return cast m;
        });
        registerMessageType(bundles.xlsx.bodies.XlsxSheet, () -> {
            trace(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> CREATING XLSX SHEET");
            var m = @:privateAccess new Message<bundles.xlsx.bodies.XlsxSheet>();
            m.body = new bundles.xlsx.bodies.XlsxSheet({});
            return cast m;
        });

        registerBodyConverter(bundles.xlsx.bodies.XlsxSheet, CsvBody, (sheet) -> {
            trace(">>>>>>>>>>>>>>>>>>>>>>> DOING CONVERSION FROM XLSXSHEET -> CSVBODY");
            var firstRow = null;
            var csv = new CsvBody();
            for (row in sheet.rows()) {
                if (firstRow == null) {
                    firstRow = row;
                    csv.addColumns(firstRow);
                } else {
                    csv.addRow(row);
                }
            }
            return csv;
        });
    }
}