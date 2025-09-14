import { Person, personDef, Requisition, requisitionDef } from "../util/data/DBSchema";
import { TableRepository, TypedTable, WorkbookORM } from "../util/data/UniversalRepo";

export class RequisitionService {
    constructor (private requisitionRepo: TypedTable<Requisition>, private personRepo: TypedTable<Person>) {
    }

    public static async create(ctx : Excel.RequestContext) : Promise<RequisitionService> {
        const orm = new WorkbookORM(ctx.workbook);

        const requisitionRepo = await orm.tables.getAs<Requisition>("Requisition", requisitionDef);
        const personRepo = await orm.tables.getAs<Person>("Person", personDef);

        const service = new RequisitionService(requisitionRepo, personRepo);
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

    // public async getAllDraftSheetNames() : Promise<string[]> {
    //     return 
    // }

}