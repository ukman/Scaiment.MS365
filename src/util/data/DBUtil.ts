import { TableDefinition } from "./UniversalRepo";

export function coerceAnyToDBType<T extends Record<string, any>>(dto : any, tableDefinition : TableDefinition<T>) {
    const keys = Object.keys(tableDefinition.columns);
    for(let i = 0; i < keys.length; i++) {
        const colName = keys[i];
        const col = tableDefinition.columns[colName];
        console.log("col = ", colName, col);
        let curValue = dto[colName];
        console.log("curValue = ", curValue);
        console.log("typeof curValue = ", typeof(curValue));
        if(col.required) {
            if(typeof(curValue) == 'undefined') {
                curValue = col.default;
            }            
        }
        if(col.type == 'number' && typeof(curValue) != 'number') {
            curValue = +curValue;            
        }
        if(col.type == 'string' && typeof(curValue) != 'string') {
            if(typeof(curValue) != 'undefined') {
                curValue = "" + curValue;            
            }
        }
        if(typeof(curValue) != 'undefined')
            dto[colName] = curValue;
        console.log("dto[colName] = ", dto[colName]);
        console.log("typeof(dto[colName]) = ", typeof(dto[colName]));
    }
    return dto;

}