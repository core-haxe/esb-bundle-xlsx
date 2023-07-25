package bundles.xlsx;

import esb.core.Bus.*;
import esb.core.bodies.CsvBody;
import esb.core.Message;

class Bundle extends esb.core.Bundle {
    public override function start() {
        super.start();

        registerMessageType(bundles.xlsx.bodies.XlsxBody, () -> {
            var m = @:privateAccess new Message<bundles.xlsx.bodies.XlsxBody>();
            m.body = new bundles.xlsx.bodies.XlsxBody();
            return cast m;
        });

        registerMessageType(bundles.xlsx.bodies.XlsxSheet, () -> {
            var m = @:privateAccess new Message<bundles.xlsx.bodies.XlsxSheet>();
            m.body = new bundles.xlsx.bodies.XlsxSheet({});
            return cast m;
        });

        registerBodyConverter(bundles.xlsx.bodies.XlsxSheet, CsvBody, (sheet) -> {
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