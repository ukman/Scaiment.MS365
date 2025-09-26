import { Person, personDef, Requisition, requisitionDef, RequisitionItem, requisitionItemDef, RequisitionApproval, requisitionApprovalDef, ProjectMember } from "../util/data/DBSchema";
import { SchemaValidator, TableRepository, TypedTable, WorkbookORM } from "../util/data/UniversalRepo";
import { excelLog } from "../util/Logs";
import { ExcelService } from "./ExcelService";
import { PersonService } from "./PersonService";
import { ProcurementRole, ProjectService } from "./ProjectService";

export class RequisitionService {
    constructor (
        private context : Excel.RequestContext,
        private requisitionRepo: TypedTable<Requisition>, 
        private requisitionItemRepo: TypedTable<RequisitionItem>, 
        private requisitionApprovalRepo: TypedTable<RequisitionApproval>, 
        private personRepo: TypedTable<Person>) {
    }

    public static async create(context : Excel.RequestContext) : Promise<RequisitionService> {
        const orm = new WorkbookORM(context.workbook);

        const requisitionRepo = await orm.tables.getAs<Requisition>("Requisition", requisitionDef);
        const requisitionItemRepo = await orm.tables.getAs<RequisitionItem>("RequisitionItem", requisitionItemDef);
        const requisitionApprovalRepo = await orm.tables.getAs<RequisitionApproval>("RequisitionApproval", requisitionApprovalDef);
        const personRepo = await orm.tables.getAs<Person>("Person", personDef);

        const service = new RequisitionService(context, requisitionRepo, requisitionItemRepo, requisitionApprovalRepo, personRepo);
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

    public async findAllByEmailIds(emailIds : string[]) : Promise<Requisition[]> {
        const res = await this.requisitionRepo.rows.findAllByKeys("emailId", emailIds);
        return res.map(item => item.row);     
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

}