import * as React from "react";
import { useState } from 'react';
import {
  Stack,
  Text,
  SearchBox,
  DefaultButton,
  PrimaryButton,
  Dropdown,
  IDropdownOption,
  Pivot,
  PivotItem,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  MessageBar,
  getTheme,
  IColumn,
} from "@fluentui/react";

import { WorkbookSchemaGenerator } from "../../util/data/SchemaGenerator";
import { WorkbookORM, TableRepository, TypedRowRepository, SchemaValidator } from "../../util/data/UniversalRepo";
import { ExcelService } from "../../services/ExcelService";
import { coerceAnyToDBType } from "../../util/data/DBUtil";
import { Project, ProjectMember, Requisition, requisitionDef, RequisitionItem } from "../../util/data/DBSchema";
import { excelLog } from "../../util/Logs";
import { RequisitionService } from "../../services/RequisitionService";
import { PersonService } from "../../services/PersonService";
import { ProcurementRole, ProjectService } from "../../services/ProjectService";

export interface TestProps {
    title: string;
}
  
const Test: React.FC<TestProps> = (props: TestProps) => {

    const formatDate = (iso: Date) => new Date(iso).toLocaleDateString();
 
    const { title } = props;
    const [ data, setData] = useState<string>('---');

    const [items, setItems] = useState<Requisition[]>([]);

    const testCoerce = async () => {
        console.log("Test coerce");
        const coerced = coerceAnyToDBType({}, requisitionDef);

        console.log("Coerced : ", coerced);
    }

    const addDraft = async () => {
        try {
            await Excel.run(async (context) => {

                const personService = await PersonService.create(context);
                excelLog("Trying to get current user ");
                const currentUser = await personService.getCurrentUser();
                excelLog("currentUser = " + currentUser);

                const projectService = await ProjectService.create(context);
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

                const requisition : Requisition = {
                    // createdBy : 
                    createdAt : new Date()

                } as Requisition;
                // Получаем исходный лист
                const sourceSheet = context.workbook.worksheets.getItem("RequisitionTemplate");

                const excelService = await ExcelService.create(context);
                
                // Копируем лист с новым именем
                const newName = "Draft " + Math.floor(Math.random() * 100);
                const copiedSheet = sourceSheet.copy(Excel.WorksheetPositionType.end);
                copiedSheet.name = newName;
              
                // Получаем именованную ячейку "createdAt" на новом листе
                const namedItem = copiedSheet.names.getItemOrNullObject("createdAt");              

                // Синхронизируем изменения
                await context.sync();

                const newRequisition : Requisition = {
                    createdAt : SchemaValidator.toExcelValue("date", new Date()),
                    createdBy : currentUser.id,
                    createdByName : currentUser.fullName,
                    projectName: (authorProjects.length > 0 ? authorProjects[0].projectName : ""),                   
                    responsibleName : responsibles.length == 1 ? responsibles[0] : "",
                    
                } as Requisition;
                approvers.forEach((a, i) => {
                    (newRequisition as any)["approver" + i] = a.personName;
                });

                excelService.fillSheet(newRequisition, newName);

                // Синхронизируем изменения
                await context.sync();


                /*
                if (!namedItem.isNullObject) {
                    // Получаем диапазон и устанавливаем дату/время
                    const range = namedItem.getRange();

                    const currentDateTime = SchemaValidator.toExcelValue("date", new Date());

                    range.values = [[currentDateTime]];
                }                
                // console.log(`Лист "${sourceSheetName}" успешно скопирован с новым именем "${newSheetName}".`);
                */
            });
          } catch (error) {
            excelLog("Ошибка при копировании листа: " + error + "\n" + JSON.stringify(error));
          }        
    }

    const addRequisition = async () => {
        await Excel.run(async (ctx) => {
            try {
                await excelLog("Step1");
                const requisitionService = await RequisitionService.create(ctx);
                const excelService = await ExcelService.create(ctx);
                const activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
                activeSheet.load("name");
                await ctx.sync();
                
                const requisition = await excelService.getNamedRangesWithValues(activeSheet.name);
                const items = await excelService.getTablesDataFromSheet(activeSheet.name);
                for(const key in items) {
                    const value = items[key];
                    if(Array.isArray(value)) {
                        (requisition as any).RequisitionItems = value as RequisitionItem[];
                    }
                }

                await excelLog("Step5 " + JSON.stringify(requisition));
                await requisitionService.saveRequisition(requisition as Requisition);


//                await reqRepo.rows.add(data);
                await excelLog("Saved");
                await ctx.sync();
                setData("Req = " + JSON.stringify(data));
            } catch (e) {
                setData("Error " + e + " " + JSON.stringify(e));
                await excelLog("Error = " + e + "\n" + JSON.stringify(e));
            } finally {
                console.log("Finally");
            }
        });

    }

    const getData = async () => {
        await Excel.run(async (ctx) => {
            try {
                const orm = new WorkbookORM(ctx.workbook);
                const reqRepo = await orm.tables.getAs<Requisition>("Requisition", requisitionDef);
                await excelLog(JSON.stringify(reqRepo, null, 2));
        
                const excelService = await ExcelService.create(ctx);
                const data = await excelService.getNamedRangesWithValues("RequisitionTemplate");
                const items = await excelService.getTablesDataFromSheet("RequisitionTemplate");
                data.items = items;
                setData(JSON.stringify(data, null, 2));
            } catch (e) {
                setData(JSON.stringify(e));
            } finally {
                console.log("Finally");
            }
        });
    }
    const getDrafts = async () => {
        
        await Excel.run(async (ctx) => {
            try {
                const requisitionService = await RequisitionService.create(ctx);
                const drafts = await requisitionService.getDrafts(ctx);
                // setData(JSON.stringify(drafts, null, 2));
                setItems(drafts); 
            } catch (e) {
                setData(JSON.stringify(e));
                throw e;
            } finally {
                console.log("Finally");
            }
        });
    }

    interface ChatCompletionRequest {
        messages: Array<{ role: 'system' | 'user' | 'assistant'; content: string }>;
        model: string;
        max_tokens?: number;
        temperature?: number;
    }
    
    interface ChatCompletionResponse {
        choices: Array<{ message: { content: string } }>;
        id: string;
        model: string;
        usage: { prompt_tokens: number; completion_tokens: number; total_tokens: number };
    }
    
    const testRestApi = async () => {
        await excelLog("Start test rest api");   
        const requestBody: ChatCompletionRequest = {
            messages:[{role: "user", content: "Назови первые 10 простых чисел"}],
            model:"gpt-4.1",
            temperature:0.1,
            max_tokens:100
        };
    
        try {
//            const response = await fetch("https://localhost:3000/taskpane.html", {
            const response = await fetch("https://localhost:3001/api/openai/chat", {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(requestBody),
            });  
            const txt = await response.text();
            await excelLog("Response status = " + response.status + " text = " + txt);      
        }catch(e) {
            await excelLog("Error = " + e);      
        }
    }


    const columns : IColumn[] = [
        { key: "id", name: "Request", minWidth: 100, maxWidth:100, onRender: (i: Requisition) => (
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="mediumPlus"><b>{i.id}</b></Text>
            <Text variant="xSmall">{i.requisitionItems} items</Text>
          </Stack>
        )},
        { key: "proj", name: "Name / Project", minWidth: 150, onRender: (i: Requisition) => (
          <Stack tokens={{ childrenGap: 4 }}>
            <Text>{i.name}</Text>
            <Text variant="xSmall">{i.projectName}</Text>
          </Stack>
        )},
        { key: "needed", name: "Needed by", minWidth: 100, onRender: (i: Requisition) => (
          <Stack tokens={{ childrenGap: 4 }}>
            <Text>{formatDate(i.dueDate)}</Text>
            <Text variant="xSmall">Created {formatDate(i.createdAt)}</Text>
          </Stack>
        )},
        // { key: "status", name: "Status", minWidth: 100, onRender: (i: Requisition) => (
        //   <Stack horizontal tokens={{ childrenGap: 8 }}>
        //     <Chip text={i.status} color={statusColorMap[i.status]} />
        //     {/* <Chip text={i.risk} color={riskColorMap[i.risk]} /> */}
        //   </Stack>
        // )},
        // { key: "total", name: "Total", minWidth: 100, onRender: (i: RequestRow) => <Text>{currency(i.total)}</Text> },
        { key: "owner", name: "Owner / Updated", minWidth: 120, onRender: (i: Requisition) => (
          <Stack>
            <Text>{i.responsibleName}</Text>
            {/* <Text variant="xSmall">{formatDate(i.lastUpdate)}</Text> */}
          </Stack>
        )},
        /*
        { key: "actions", name: "Actions", minWidth: 180, onRender: (i: RequestRow) => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            {i.status === "Awaiting Approval" && (
              <PrimaryButton text="Approve" onClick={() => alert(`Approved ${i.id}`)} />
            )}
            {i.status === "Approved" && (
              <DefaultButton text="Create PO" onClick={() => alert(`Create PO for ${i.id}`)} />
            )}
            <DefaultButton text="Received" onClick={() => alert(`Received ${i.id}`)} />
          </Stack>
        )},
        */ 
      ];
    
    
    return (
        <div>
            Test {title}
            <PrimaryButton
                onClick={getData}
                className="w-100"
            >
                Get Data2
            </PrimaryButton>
            <PrimaryButton
                onClick={testCoerce}
                className="w-100"
            >
                Coerce
            </PrimaryButton>
            <PrimaryButton
                onClick={addRequisition}
                className="w-100"
            >
                Add Requisition
            </PrimaryButton>
            <PrimaryButton
                onClick={addDraft}
                className="w-100"
            >
                Add Requisition Draft
            </PrimaryButton>
            <PrimaryButton
                onClick={getDrafts}
                className="w-100"
            >
                Get Drafts
            </PrimaryButton>
            <PrimaryButton
                onClick={testRestApi}
                className="w-100"
            >
                Test Rest API
            </PrimaryButton>
<pre><code>{data}</code></pre>

    <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
        <Stack grow styles={{ root: { minWidth: 500 } }}>
          <DetailsList
            items={items}
            columns={columns as any}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
          />
        </Stack>
    </Stack>
        </div>

      );
    
}

export default Test;
