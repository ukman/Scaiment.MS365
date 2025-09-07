import { Person, personDef } from "../util/data/DBSchema";
import { TableRepository, TypedTable, WorkbookORM } from "../util/data/UniversalRepo";

export class PersonService {
    constructor (private personRepo: TypedTable<Person>) {
    }

    public static async create(ctx : Excel.RequestContext) : Promise<PersonService> {
        const orm = new WorkbookORM(ctx.workbook);

        const personRepo = await orm.tables.getAs<Person>("Person", personDef);
        personRepo.rows.findAllBy("id", 109);

        const service = new PersonService(personRepo);
        return service;
    }

    /*
    public findPerson(id : number) : Person {
        return this.personRepo.findById(id);
    }
        */
}