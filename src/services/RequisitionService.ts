import { Person, personDef } from "../util/data/DBSchema";
import { TableRepository, TypedTable, WorkbookORM } from "../util/data/UniversalRepo";

export class RequisitionService {
    constructor (personRepo: TypedTable<Person>) {
    }

    public static async create(ctx : Excel.RequestContext) : Promise<RequisitionService> {
        const orm = new WorkbookORM(ctx.workbook);

        const personRepo = await orm.tables.getAs<Person>("Person", personDef);

        const service = new RequisitionService(personRepo);
        return service;
    }

}