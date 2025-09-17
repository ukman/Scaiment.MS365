import { Person, personDef } from "../util/data/DBSchema";
import { TableRepository, TypedTable, WorkbookORM } from "../util/data/UniversalRepo";

export class PersonService {
    constructor (private personRepo: TypedTable<Person>) {
    }

    public static async create(ctx : Excel.RequestContext) : Promise<PersonService> {
        const orm = new WorkbookORM(ctx.workbook);

        const personRepo = await orm.tables.getAs<Person>("Person", personDef);

        const service = new PersonService(personRepo);
        return service;
    }

    public async getCurrentUser() : Promise<Person> {
        // TODO add real current user 
        return this.findPersonById(1);
    }

    public async findPersonById(id : number) : Promise<Person> {
        const person = this.personRepo.rows.findFirstBy("id", id);
        return (await person).row;
    }
}