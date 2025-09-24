import { ColumnDefinition, GeneratedTable, TableDefinition, WorkbookSchemaGenerator } from "../util/data/SchemaGenerator";
import { WorkbookORM } from "../util/data/UniversalRepo";
import { excelLog } from "../util/Logs";

export class MetadataService {
    constructor (private context : Excel.RequestContext) {

    }

    public static async create(ctx : Excel.RequestContext) : Promise<MetadataService> {
        const service = new MetadataService(ctx);
        return service;
    }

    public async generateRequisitionSystemMessage() : Promise<string> {
        const schema = await this.generateSchema();
        const reqSchema = schema.runtime["Requisition"]; 
        const reqItemSchema = schema.runtime["RequisitionItem"]; 

        const reqSM = this.generateSystemMessageForColumns(reqSchema.columns, "  ");
        const reqItemSM = this.generateSystemMessageForColumns(reqItemSchema.columns, "    ");

        return `Ты- AI помощник менеджера по снабжению в строительной компании.
В отдел снабжения приходят письма по email с заявками на разные строительные проекты. 
Твоя задача- проанализировать письмо и создать JSON с заявкой. 
Этот JSON будет сохранен в базе данных и использован в системе управления заявками.
JSON документ состоит из полей:
${reqSM}
  - 'requisitionItems' list of products
${reqItemSM}`;
    }

    private generateSystemMessageForColumns(columns : Record<string, ColumnDefinition<any>>, prefix : string) : string {
        const res = [];
        for(const col in columns) {            
            const colDesc = columns[col];
            if(!colDesc.calculated && colDesc.aiDescription && colDesc.aiDescription.trim().length > 0) {
                res.push(`${prefix}- '${col}' ${colDesc.aiDescription} `);
            }
        }
        return res.join("\n");                
    }

    public async generateSchema() : Promise<{ code: string; runtime: Record<string, TableDefinition<Record<string, any>>>; tables: GeneratedTable[] }> {
        const orm = new WorkbookORM(this.context.workbook);
        const gen = new WorkbookSchemaGenerator(this.context.workbook);
        const out = await gen.generateTypeScript({
            sampleRows: 50,
            emitHeader: true,
            inlineTableDefinition: false, // если true — без импортов
            importPath: "./UniversalRepo",
        });
        return out;
    }    
}