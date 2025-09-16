import * as React from "react";
import { useState } from 'react';
import { PrimaryButton } from '@fluentui/react';
import { WorkbookSchemaGenerator } from "../../util/data/SchemaGenerator";
import { WorkbookORM, TableRepository, TypedRowRepository } from "../../util/data/UniversalRepo";
import { ExcelService } from "../../services/ExcelService";
import { coerceAnyToDBType } from "../../util/data/DBUtil";
import { Requisition, requisitionDef } from "../../util/data/DBSchema";
import { excelLog } from "../../util/Logs";
import { RequisitionService } from "../../services/RequisitionService";


export interface TestProps {
    title: string;
}
  
const Test: React.FC<TestProps> = (props: TestProps) => {
    const { title } = props;
    const [ data, setData] = useState<string>('---');

    const testCoerce = async () => {
        console.log("Test coerce");
        const coerced = coerceAnyToDBType({}, requisitionDef);

        console.log("Coerced : ", coerced);
    }

    const addRequisition = async () => {
        await Excel.run(async (ctx) => {
            try {
                await excelLog("Step1");
                const orm = new WorkbookORM(ctx.workbook);
                const requisitionService = await RequisitionService.create(ctx);
                const reqRepo = await orm.tables.getAs<Requisition>("Requisition", requisitionDef);
                const excelService = await ExcelService.create(ctx);
                const requisitionSimple = await excelService.getNamedRangesWithValues("RequisitionTemplate");
                const items = await excelService.getTablesDataFromSheet("RequisitionTemplate");
                const requisition = { ...requisitionSimple, ...items};

                await excelLog("Step5 " + JSON.stringify(requisition));
                await requisitionService.saveRequisition(requisition as Requisition);


//                await reqRepo.rows.add(data);
                await excelLog("Saved");
                await ctx.sync();
                setData("Req = " + JSON.stringify(data));
            } catch (e) {
                setData("Error " + e + " " + JSON.stringify(e));
                await excelLog("Error = " + e + "\n" + JSON.stringify(e));
            } finally {
                console.log("Finally");
            }
        });

    }

    const getData = async () => {
        await Excel.run(async (ctx) => {
            try {
                const orm = new WorkbookORM(ctx.workbook);
                const reqRepo = await orm.tables.getAs<Requisition>("Requisition", requisitionDef);
                await excelLog(JSON.stringify(reqRepo, null, 2));
        
                const excelService = await ExcelService.create(ctx);
                const data = await excelService.getNamedRangesWithValues("RequisitionTemplate");
                const items = await excelService.getTablesDataFromSheet("RequisitionTemplate");
                data.items = items;
                setData(JSON.stringify(data, null, 2));
            } catch (e) {
                setData(JSON.stringify(e));
            } finally {
                console.log("Finally");
            }
        });
    }

    
    return (
        <div>
            Test {title}
            <PrimaryButton
                onClick={getData}
                className="w-100"
            >
                Get Data2
            </PrimaryButton>
            <PrimaryButton
                onClick={testCoerce}
                className="w-100"
            >
                Coerce
            </PrimaryButton>
            <PrimaryButton
                onClick={addRequisition}
                className="w-100"
            >
                Add Requisition
            </PrimaryButton>
<pre><code>{data}</code></pre>
        </div>
      );
    
}

export default Test;
