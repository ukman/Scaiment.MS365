import { ProjectMember, Requisition } from "../util/data/DBSchema";
import { SchemaValidator } from "../util/data/UniversalRepo";
import { excelLog } from "../util/Logs";
import { ExcelService } from "./ExcelService";
import { PersonService } from "./PersonService";
import { ProcurementRole, ProjectService } from "./ProjectService";

export class DraftService {
    constructor (private context : Excel.RequestContext) {

    }

    public static async create(ctx : Excel.RequestContext) : Promise<DraftService> {
        const service = new DraftService(ctx);
        return service;
    }

    public async addRequisitionDraft(requisitionDraft : Requisition) {
        const personService = await PersonService.create(this.context);
        // await excelLog("Trying to get current user ");
        const currentUser = await personService.getCurrentUser();
        // await excelLog("currentUser = " + currentUser);

        const projectService = await ProjectService.create(this.context);
        const authorProjects = await projectService.getUserProjectsByRole(currentUser.id, "requisition_author");
        // const firstProject : Project = (authorProjects.length == 1 ? await projectService.findProjectById(authorProjects[0].projectId) : {} as Project);
        const approvers = authorProjects.length == 1 ? await projectService.getProjectMembersByRole(authorProjects[0].projectId, "requisition_approver") : [] as ProjectMember[];
        const responsibles = [];
        const allProjectMembers = await projectService.getAllProjectMembers();
        const pmRole : ProcurementRole = "procurement_manager";
        authorProjects.forEach(p => {
            const managers = allProjectMembers.filter(pm => (pm.roleName == pmRole && pm.projectId == p.projectId));
            excelLog("managers = " + JSON.stringify(managers) + " for " + p.projectName);
            managers.forEach(pm => {
                if(!responsibles.includes(pm.personName))
                    responsibles.push(pm.personName);
            });
        });
        excelLog("responsibles = " + JSON.stringify(responsibles));
        //excelLog("responsibles = " + JSON.stringify(responsibles));

        const newRequisition : Requisition = requisitionDraft;
        newRequisition.createdAt = new Date();

        // Получаем исходный лист
        const sourceSheet = this.context.workbook.worksheets.getItem("RequisitionTemplate");

        const excelService = await ExcelService.create(this.context);
        
        // Копируем лист с новым именем
        const newName = "Draft " + Math.floor(Math.random() * 100);
        const copiedSheet = sourceSheet.copy(Excel.WorksheetPositionType.end);
        copiedSheet.name = newName;
      
        // Получаем именованную ячейку "createdAt" на новом листе
        const namedItem = copiedSheet.names.getItemOrNullObject("createdAt");              

        // Синхронизируем изменения
        await this.context.sync();

        newRequisition.createdAt = SchemaValidator.toExcelValue("date", new Date());
        newRequisition.createdBy = currentUser.id;
        newRequisition.createdByName = currentUser.fullName;
        newRequisition.projectName = (authorProjects.length > 0 ? authorProjects[0].projectName : "");
        newRequisition.responsibleName = responsibles.length == 1 ? responsibles[0] : "";
        
        approvers.forEach((a, i) => {
            (newRequisition as any)["approver" + i] = a.personName;
        });

        excelService.fillSheet(newRequisition, newName);
        excelService.fillTableWithData(newName, (newRequisition as any).requisitionItems as any[]);

        // Синхронизируем изменения
        await this.context.sync();


        /*
        if (!namedItem.isNullObject) {
            // Получаем диапазон и устанавливаем дату/время
            const range = namedItem.getRange();

            const currentDateTime = SchemaValidator.toExcelValue("date", new Date());

            range.values = [[currentDateTime]];
        }                
        // console.log(`Лист "${sourceSheetName}" успешно скопирован с новым именем "${newSheetName}".`);
        */


}

}