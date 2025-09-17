import { Project, projectDef, ProjectMember, projectMemberDef } from "../util/data/DBSchema";
import { RowMatchTyped, TypedTable, WorkbookORM } from "../util/data/UniversalRepo";

export type ProcurementRole = "procurement_manager" | "requisition_author" | "requisition_approver";

export class ProjectService {
    constructor (private projectRepo: TypedTable<Project>,
        private projectMemberRepo: TypedTable<ProjectMember>,
    ) {
    }


    public static async create(ctx : Excel.RequestContext) : Promise<ProjectService> {
        const orm = new WorkbookORM(ctx.workbook);

        const projectRepo = await orm.tables.getAs<Project>("Project", projectDef);
        const projectMemberRepo = await orm.tables.getAs<ProjectMember>("ProjectMember", projectMemberDef);

        const service = new ProjectService(projectRepo, projectMemberRepo);
        return service;
    }

    public async findProjectById(projectId : number) : Promise<Project> {
        const project = await this.projectRepo.rows.findFirstBy("id", projectId);
        return project.row;
    }

    public async getUserProjectsByRole(personId : number, role : ProcurementRole) : Promise<ProjectMember[]> {
        const projectMembers = await this.getUserProjects(personId);
        return projectMembers.filter(pm => pm.roleName === role);
    }

    public async getUserProjects(personId : number) : Promise<ProjectMember[]> {
        const projectMembers = await this.projectMemberRepo.rows.findAllBy("personId", personId);
        return projectMembers.map(m => m.row);
    }

    public async getProjectMembers(projectId : number) : Promise<ProjectMember[]> {
        const projectMembers = await this.projectMemberRepo.rows.findAllBy("projectId", projectId);
        return projectMembers.map(m => m.row);
    }

    public async getProjectMembersByRole(projectId : number, role : ProcurementRole) : Promise<ProjectMember[]> {
        const projectMembers = await this.getProjectMembers(projectId);
        return projectMembers.filter(m => m.roleName === role);
    }

    public async getAllProjectMembers() : Promise<ProjectMember[]> {
        const projectMembers = await this.projectMemberRepo.rows.getAll();
        return projectMembers;
    }
}