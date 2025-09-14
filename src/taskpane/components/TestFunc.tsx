import * as React from "react";
import { useState } from 'react';
import { PrimaryButton } from '@fluentui/react';
import { WorkbookSchemaGenerator } from "../../util/data/SchemaGenerator";
import { WorkbookORM } from "../../util/data/UniversalRepo";
import { ExcelService } from "../../services/ExcelService";



export interface TestProps {
    title: string;
}
  
const Test: React.FC<TestProps> = (props: TestProps) => {
    const { title } = props;
    const [ data, setData] = useState<string>('---');

    const getData = async () => {
        await Excel.run(async (ctx) => {
            try {
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
                Get Data
            </PrimaryButton>
<pre><code>{data}</code></pre>
        </div>
      );
    
}

export default Test;
