import { Person, personDef } from "../util/data/DBSchema";
import { TableRepository, TypedTable, WorkbookORM } from "../util/data/UniversalRepo";
import { excelLog } from "../util/Logs";
import { MSGraphService } from "./MSGraphService";

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
        if(!MSGraphService.getInstance().isAuthenticated()) {
            throw new Error("You are not authenticated. Cannot get current user.");
        }
        const currentUser = MSGraphService.getInstance().getCurrentUser();
        const emails = [];
        if(currentUser.mail) {
            emails.push(currentUser.mail);
        }
        if(currentUser.userPrincipalName) {
            emails.push(currentUser.userPrincipalName);
        }
        excelLog("Trying to find user by " + JSON.stringify(emails));
        for(let i = 0; i < emails.length; i++) {
            const email = emails[i]; 
            const res = await this.findPersonByEmail(email);
            excelLog("email = " + email + " res = " + res);
            if(res) {
                return res;
            }
        }
        throw new Error("Cannot find current user by emails " + JSON.stringify(emails));
    }

    public async findPersonById(id : number) : Promise<Person> {
        const person = this.personRepo.rows.findFirstBy("id", id);
        return (await person).row;
    }

    public async findPersonByEmail(email : string) : Promise<Person> {
        const lowerEmail = email.toLowerCase();
        const person = await this.personRepo.rows.findFirstBy("email", lowerEmail);
        if(person)
            return person.row;
        return undefined;
    }
}