import * as React from "react";
import { useState } from 'react';
import { PrimaryButton } from '@fluentui/react';
import { WorkbookSchemaGenerator } from "../../util/data/SchemaGenerator";
import { WorkbookORM } from "../../util/data/UniversalRepo";
import { MetadataService } from "../../services/MetadataService";
import { excelLog } from "../../util/Logs";



export interface DraftsProps {
    title: string;
}
  
const Drafts: React.FC<DraftsProps> = (props: DraftsProps) => {
    const { title } = props;
    const [schema, setSchema2] = useState<string>('');

    async function generateSM() {
        await Excel.run(async (ctx) => {      
            try {      
            const service = await MetadataService.create(ctx);
            const sm = await service.generateRequisitionSystemMessage();
            setSchema2(sm);
            } catch(e) {
                excelLog("Error : " + e);
            }
        });
    }

    async function generateSchema() {
        await Excel.run(async (ctx) => {            
            const orm = new WorkbookORM(ctx.workbook);
            const gen = new WorkbookSchemaGenerator(ctx.workbook);
            const out = await gen.generateTypeScript({
                sampleRows: 50,
                emitHeader: true,
                inlineTableDefinition: false, // если true — без импортов
                importPath: "./UniversalRepo",
            });
            setSchema2(out.code);

        });
    }

    
    return (
        <div>
            Drafts {title}
            <PrimaryButton
                // variant="primary"
                onClick={generateSchema}
                className="w-100"
            >
                Generate Schema
            </PrimaryButton>
            <PrimaryButton
                // variant="primary"
                onClick={generateSM}
                className="w-100"
            >
                Generate SM
            </PrimaryButton>
<pre><code>{schema}</code></pre>
        </div>
      );
    
}

export default Drafts;
