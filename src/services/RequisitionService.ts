import { Person, personDef, Requisition, requisitionDef, RequisitionItem, requisitionItemDef, RequisitionApproval, requisitionApprovalDef } from "../util/data/DBSchema";
import { TableRepository, TypedTable, WorkbookORM } from "../util/data/UniversalRepo";
import { excelLog } from "../util/Logs";
import { ExcelService } from "./ExcelService";

export class RequisitionService {
    constructor (
        private requisitionRepo: TypedTable<Requisition>, 
        private requisitionItemRepo: TypedTable<RequisitionItem>, 
        private requisitionApprovalRepo: TypedTable<RequisitionApproval>, 
        private personRepo: TypedTable<Person>) {
    }

    public static async create(ctx : Excel.RequestContext) : Promise<RequisitionService> {
        const orm = new WorkbookORM(ctx.workbook);

        const requisitionRepo = await orm.tables.getAs<Requisition>("Requisition", requisitionDef);
        const requisitionItemRepo = await orm.tables.getAs<RequisitionItem>("RequisitionItem", requisitionItemDef);
        const requisitionApprovalRepo = await orm.tables.getAs<RequisitionApproval>("RequisitionApproval", requisitionApprovalDef);
        const personRepo = await orm.tables.getAs<Person>("Person", personDef);

        const service = new RequisitionService(requisitionRepo, requisitionItemRepo, requisitionApprovalRepo, personRepo);
        return service;
    }

    public async findById(id : number) : Promise<Requisition> {
        const res = await this.requisitionRepo.rows.findFirstBy("id", id);
        return res.row;     
    }

    public async findAll() : Promise<Requisition[]> {
        const res = await this.requisitionRepo.rows.getAll();        
        return res;     
    }

    public async saveRequisition(requisition : Requisition) {

        requisition.createdAt = new Date();
        await this.requisitionRepo.rows.add(requisition);
        ((requisition as any).RequisitionItems as RequisitionItem[]).forEach(ri => ri.requisitionId = requisition.id);
        await this.requisitionItemRepo.rows.addMany((requisition as any).RequisitionItems);

        const approvals = (requisition.approvals ? requisition.approvals.split(",") : [])
            .map(s => +s)
            .filter(n => n > 0)
            .map(personId => {
                const ra : RequisitionApproval = {
                    requisitionId: requisition.id,
                    createdBy: requisition.createdBy,
                    personId: personId,
                    decision: 0
                } as RequisitionApproval;
                return ra;
            });
        await excelLog("Approvals = " + JSON.stringify(approvals))
        await this.requisitionApprovalRepo.rows.addMany(approvals);

    }

    // public async getAllDraftSheetNames() : Promise<string[]> {
    //     return 
    // }

    public async getDrafts(ctx : Excel.RequestContext) : Promise<Requisition[]> {
        const excelService = await ExcelService.create(ctx);
        const draftSheetNames = await excelService.findSheetsWithMarker(ctx, "__requisitionDraftMarker");

        const drafts : Requisition[] = [];
        for(let i = 0; i < draftSheetNames.length; i++) {
            const sheetName = draftSheetNames[i];
            const draftData = (await excelService.getNamedRangesWithValues(sheetName)) as Requisition;
            
            const items = await excelService.getTablesDataFromSheet(sheetName);
            for(const key in items) {
                if(key.startsWith("requisitionItems")) {
                    const value = items[key];
                    if(Array.isArray(value)) {
                        (draftData as any).RequisitionItems = value as RequisitionItem[];
                    }
                } 
            }
            
            drafts.push(draftData);
        }
        return drafts;
    }


}