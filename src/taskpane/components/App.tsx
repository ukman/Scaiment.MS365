import * as React from 'react';
import { useState } from 'react';
import { useAuth } from '../../auth/AuthProvider';
import { Button, Alert, Spinner } from 'react-bootstrap';
import { TableDefinition, WorkbookORM } from '../../util/data/UniversalRepo';
import { WorkbookSchemaGenerator } from '../../util/data/SchemaGenerator';
import { Person, personDef } from '../../util/data/DBSchema';


interface AppProps {
  title: string;
}

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
  names: { /* Email: "E-mail" */ },
  // Optional: control column order when writing (defaults to header order)
  order: ["Id", "FirstName", "LastName", "FullName", "Email", "IsActive", "CreatedAt"],
};


const App: React.FC<AppProps> = ({ title }) => {
  const { getToken } = useAuth();
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string>('');
  const [users, setUsers] = useState<User[]>([]);

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

  
  const fetchTables = async () => {
    setIsLoading(true);
    setError('');
    try {
      // Insert into Excel
      await Excel.run(async (ctx) => {
        const orm = new WorkbookORM(ctx.workbook);

        const users = await orm.tables.getAs<Person>("Person", personDef);
        // const p222 : Person = {} as any;
    
        // Read typed records
        // const t0 = new Date().getTime();
        // const list: User[] = await users.rows.getAll();    
        // const t = (new Date()).getTime() - t0;
        // setUsers(list);



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

  return (
    <div className="container p-3">
      <h1 className="text-center mb-4">{title}</h1>
      {error && <Alert variant="danger">{error}</Alert>}
      {isLoading && <Spinner animation="border" className="d-block mx-auto" />}
      <Button
        variant="primary"
        onClick={fetchLatestEmail}
        disabled={isLoading}
        className="w-100"
      >
        {isLoading ? 'Processing...' : 'Get Latest Email'}
      </Button>
      <Button
        variant="primary"
        onClick={fetchTables}
        disabled={isLoading}
        className="w-100"
      >
        {isLoading ? 'Processing...' : 'Get Tables'}
      </Button>

      <table className="min-w-full text-sm">
        <thead className="bg-gray-50">
          <tr>
          <th>
              ID
            </th>
            <th>
              FullName
            </th>
            <th>
              Email
            </th>
            <th>
              Active
            </th>
          </tr>
        </thead>
        <tbody>
        {users.map((u) => (
          <tr key={u.Id} className="border-t">
            <td>{u.Id}</td>
            <td>{u.FullName}</td>
            <td>{u.Email ?? ""}</td>
            <td>{u.IsActive ? "Yes" : "No"}</td>
          </tr>
          ))}
        </tbody>

      </table>
    </div>
  );
};

export default App;