import { Person, personDef, Requisition, requisitionDef, RequisitionItem, requisitionItemDef, RequisitionApproval, requisitionApprovalDef } from "../util/data/DBSchema";
import { TableRepository, TypedTable, WorkbookORM } from "../util/data/UniversalRepo";
import { excelLog } from "../util/Logs";

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

    public async findById(id : number) : Promise<Person> {
        const res = this.personRepo.rows.findFirstBy("id", id);
        return (await res).row;     
    }

    public async findAll() : Promise<Requisition[]> {
        const res = this.requisitionRepo.rows.getAll();        
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

}