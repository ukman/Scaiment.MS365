import * as React from "react";
import { useState } from 'react';
import { PrimaryButton } from '@fluentui/react';
import { WorkbookSchemaGenerator } from "../../util/data/SchemaGenerator";
import { WorkbookORM } from "../../util/data/UniversalRepo";

export interface DraftsProps {
    title: string;
}
  
const Drafts: React.FC<DraftsProps> = (props: DraftsProps) => {
    const { title } = props;
    const [ mailStatus, setMailStatus] = useState<string>('---');

    const checkStatus = async () => {
        // Проверка поддержки Mailbox API
        const keys : string[] = [];
        for(const key in Office.context) {
            keys.push(key);
        }

        if (!Office.context.mailbox) {
            setMailStatus("Mailbox API not available. Ensure SSO and integrated account. keys = " + keys.join(", "));
            // console.error("Mailbox API not available. Ensure SSO and integrated account. keys = ");
            return;
        }        
        setMailStatus("Mailbox API exists. ");
    }

    
    return (
        <div>
            Mail {title}
            <PrimaryButton
                // variant="primary"
                onClick={checkStatus}
                className="w-100"
            >
                Generate Schema
            </PrimaryButton>
<pre><code>{mailStatus}</code></pre>
        </div>
      );
    
}

export default Drafts;
