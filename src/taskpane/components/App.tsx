import * as React from 'react';
import { useState } from 'react';
import { useAuth } from '../../auth/AuthProvider';
// import { Button, Alert, Spinner } from 'react-bootstrap';
import { PrimaryButton, TextField, Spinner, Stack, MessageBar } from '@fluentui/react';
import { RowMatchTyped, TableDefinition, WorkbookORM } from '../../util/data/UniversalRepo';
import { WorkbookSchemaGenerator } from '../../util/data/SchemaGenerator';
import { Person, personDef } from '../../util/data/DBSchema';
import { DetailsList, SelectionMode, IColumn, DetailsListLayoutMode, Pivot, PivotItem } from '@fluentui/react';
import Dashboard from './Dashboard';
import Pulse from './Pulse';
import Drafts from './Drafts';
import Mail from './Mail';
import OutlookEmails from './OutlookEmail';

interface AppProps {
  title: string;
}

const items = [
  { id: 1, name: 'Кабель NYM 3x1.5', quantity: 100, unit : "метр" },
  { id: 2, name: 'Подрозетник синий', quantity: 50, unit : "штука" },
  { id: 3, name: 'Розетка', quantity: 25, unit : "штука"  },
  { id: 4, name: 'Автомат 16А', quantity: 5, unit : "штука"  },
  { id: 5, name: 'Автомат 25А', quantity: 3, unit : "штука"  },
  { id: 5, name: 'УЗО 40А', quantity: 3, unit : "штука"  }
];

// Определите колонки таблицы
const columns: IColumn[] = [
  {
    key: 'column1',
    name: 'ID',
    fieldName: 'id',
    minWidth: 30,
    maxWidth: 30,
    isResizable: false,
  },
  {
    key: 'column2',
    name: 'Название',
    fieldName: 'name',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
  },
  {
    key: 'column3',
    name: 'Ед.Изм.',
    fieldName: 'unit',
    minWidth: 60,
    maxWidth: 80,
    isResizable: true,
  },
  {
    key: 'column4',
    name: 'Количество',
    fieldName: 'quantity',
    minWidth: 20,
    maxWidth: 30,
    isResizable: true,
  },
];

/*
type User = { Id: number; FirstName: string; LastName: string; FullName : string; Email: string; IsActive: boolean; CreatedAt: Date };

const userDef: TableDefinition<User> = {
  columns: {
    Id:       { type: "number",  required: true },
    FirstName:     { type: "string",  required: true },
    LastName:     { type: "string",  required: true },
    FullName:     { type: "string",  required: true },
    Email:    { type: "string" },
    IsActive: { type: "boolean", default: true },
    CreatedAt:{ type: "date",    default: () => new Date() },
  },
  // Optional: map keys to Excel header names if they differ (default = same name)
  names: { /* Email: "E-mail" * / },
  // Optional: control column order when writing (defaults to header order)
  order: ["Id", "FirstName", "LastName", "FullName", "Email", "IsActive", "CreatedAt"],
};
*/

const App: React.FC<AppProps> = ({ title }) => {
  const { getToken } = useAuth();
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string>('');
  // const [users, setUsers] = useState<Person[]>([]);
  const [users, setUsers] = useState<Person[]>([]);
  

  // Retry helper function for Graph API throttling
  const fetchWithRetry = async (url: string, options: RequestInit, retries = 3, delay = 1000) => {
    for (let i = 0; i < retries; i++) {
      try {
        const response = await fetch(url, options);
        if (response.status === 429) {
          const retryAfter = response.headers.get('Retry-After') || delay;
          console.warn(`Throttled, retrying after ${retryAfter}ms`);
          await new Promise((resolve) => setTimeout(resolve, parseInt("" + retryAfter, 10) * 1000 || delay));
          continue;
        }
        if (!response.ok) throw new Error(`HTTP ${response.status}: ${await response.text()}`);
        return response;
      } catch (err) {
        if (i === retries - 1) throw err;
        await new Promise((resolve) => setTimeout(resolve, delay * Math.pow(2, i)));
      }
    }
    throw new Error('Max retries reached');

    await Excel.run(async (ctx) => {
      const orm = new WorkbookORM(ctx.workbook);
  
      // const users = await orm.tables.getAs<User>("Users", userDef);
  
      // // Read typed records
      // const list: User[] = await users.rows.getAll();
  
      // // Add typed row with validation/coercion/defaults
      // await users.rows.add({ Id: 101, Name: "Alice", Email: "a@ex.com" });
    });    
  };

  
  const addRow = async () => {
    setIsLoading(true);
    setError('');
    try {
      // Insert into Excel
      await Excel.run(async (ctx) => {
        const orm = new WorkbookORM(ctx.workbook);

        const users = await orm.tables.getAs<Person>("Person", personDef);

        for(let j = 0; j < 3; j++) {
        const all : Partial<Person>[] = [];
        for(let i = 0; i < 100; i++) {
          const newPerson : Partial<Person> = {
            id:100 + i,
            firstName:"John" + j,
            lastName:"Smith",
            isActive: true,
            email:"john.smith@mail.com",
            createdAt: new Date()
          };
          all.push(newPerson);
        }
        await users.rows.addMany(all, {fillDefaults:true});
      }

        /*
        // const p222 : Person = {} as any;
    
        // Read typed records
        // const t0 = new Date().getTime();
        const list1: Person[] = await users.rows.getAll();    
        // setUsers(list1);
        // const t = (new Date()).getTime() - t0;
        const list2 : RowMatchTyped<Person>[] = await users.rows.findAllBy("id", 1);
        // setUsers(list2);
        const list : Person[] = list2.map(rm => rm.row);
        setUsers(list);
        // list2.map(rm => rm.)


        const gen = new WorkbookSchemaGenerator(ctx.workbook);
        const out = await gen.generateTypeScript({
          sampleRows: 50,
          emitHeader: true,
          inlineTableDefinition: false, // если true — без импортов
          importPath: "./UniversalRepo",
        });
        console.log(out.code); // скопируй это в .ts файл в проекте
        // или используй out.runtime["Users"] как TableDefinition в рантайме
        
        
        // setError("Success : " + JSON.stringify(list) + " t = " + t + " ms");  
        */
      });
    } catch (err) {
      setError(
        err.message.includes('429')
          ? 'Graph API rate limit exceeded. Please wait a moment and try again.'
          : 'Failed to fetch data: ' + err.message
      );
    } finally {
      setIsLoading(false);
    }
  };

  const fetchLatestEmail = async () => {
    setIsLoading(true);
    setError('');
    try {
      const token = await getToken();
      const response = await fetchWithRetry(
        'https://graph.microsoft.com/v1.0/me/messages?$top=1&$select=id,subject,body',
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const data = await response.json();
      const latestEmail = data.value[0] || {};

      // Insert into Excel
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange('A1').values = [[latestEmail.subject || 'No Subject']];
        sheet.getRange('A2').values = [[latestEmail.body?.content || 'No Body']];
        await context.sync();
      });
    } catch (err) {
      setError(
        err.message.includes('429')
          ? 'Graph API rate limit exceeded. Please wait a moment and try again.'
          : 'Failed to fetch email: ' + err.message
      );
    } finally {
      setIsLoading(false);
    }
  };

    const [selectedKey, setSelectedKey] = useState('item1'); // Состояние для выбранного пункта
  
    const handleLinkClick = (item) => {
      setSelectedKey(item.props.itemKey);
    };  

  return (
    <div className="container p-3 ">
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { height: '100vh', width: '100%' } }}>
      {/* Горизонтальная навигация вверху */}
      <Pivot
        aria-label="Horizontal Navigation"
        selectedKey={selectedKey}
        onLinkClick={handleLinkClick}
        styles={{
          root: { display: 'flex', justifyContent: 'center' }, // Центрирование, если нужно
          link: { minWidth: 80 }, // Минимальная ширина кнопок для 6 пунктов
        }}
      >
        <PivotItem headerText="Dashboard" itemKey="dashboard" />
        <PivotItem headerText="Pulse" itemKey="pulse" />
        <PivotItem headerText="Drafts" itemKey="drafts" />
        <PivotItem headerText="Options" itemKey="options" />
        <PivotItem headerText="Mail" itemKey="mail" />
      </Pivot>

      {/* Контент, который меняется в зависимости от выбранного пункта */}
      <div style={{ padding: 16 }}>
        {selectedKey === 'dashboard' && <Dashboard title={title}/>}
        {selectedKey === 'pulse' && <Pulse title={'Scaiment'}/>}
        {selectedKey === 'drafts' && <Drafts title={'Hello'}/>}
        {selectedKey === 'mail' && <OutlookEmails clientId='dbfefe9a-a7d6-45ce-8eee-2c3df73efe50'/>}
        {/* Добавьте аналогично для остальных */}
      </div>
    </Stack>
      
    </div>
  );
};

export default App;