import React, { useState, useEffect } from 'react';
import { Card, Button, Alert, Spinner, Badge } from 'react-bootstrap';
import { Icon, IconButton, PrimaryButton, IButtonStyles } from '@fluentui/react';
import { MSGraphService } from '../../services/MSGraphService';
import { Message } from '@microsoft/microsoft-graph-types';
import { excelLog } from '../../util/Logs';
import EmailMessageViewer from './EmailMessageViewer';
// import * as Office from '@microsoft/office-js';
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { PersonService } from '../../services/PersonService';
import { ProjectService } from '../../services/ProjectService';
import { AIService } from '../../services/AIService';
import { MetadataService } from '../../services/MetadataService';
import { Requisition } from '../../util/data/DBSchema';
import { RequisitionView } from './RequisitionView';
import { RequisitionService } from '../../services/RequisitionService';
import { DraftService, RequisitionDraft } from '../../services/DraftService';


// Инициализируем иконки (вызовите один раз в App.tsx или index.tsx)
initializeIcons();

// Интерфейсы для типизации
interface EmailData {
  id: string;
  subject: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  } | null;
  receivedDateTime: string;
  isRead: boolean;
  importance: 'low' | 'normal' | 'high';
  bodyPreview: string;
  hasAttachments: boolean;
}

interface OutlookEmailsProps {
  // Temporarily keep clientId for compatibility with App.tsx; remove later
  clientId?: string;
}

const OutlookEmails: React.FC<OutlookEmailsProps> = () => {
  const [emails, setEmails] = useState<Message[]>([]);
  const [aiResults, setAIResults] = useState<any>({});
  const [message, setMessage] = useState<Message | undefined>( undefined);
  const [requisition, setRequisition] = useState<Requisition | undefined>( undefined);
  const [draftMap, setDraftMap] = useState<Map<string, RequisitionDraft>>(new Map());
  const [requisitionMap, setRequisitionMap] = useState<Map<string, Requisition>>(new Map());
  
  const [prevMessage, setPrevMessage] = useState<Message | undefined>( undefined);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>('');

  // Автоматическая загрузка при монтировании
  useEffect(() => {
    fetchEmails();
  }, []);

  // Функция для получения токена и писем
  const fetchEmails = async () => {
    setLoading(true);
    setError('');

    try {
      let emails : string[];
      await Excel.run(async (ctx) => {            
        // const ps = await PersonService.create(ctx);
        // const ds = await DraftService.create(ctx);
        // const rs = await RequisitionService.create(ctx);
        const [ps, ds, rs] = await Promise.all([
          PersonService.create(ctx),
          DraftService.create(ctx),
          RequisitionService.create(ctx)
        ]);

        const cu = await ps.getCurrentUser();
        const projServ = await ProjectService.create(ctx);
        const projectMembers = await projServ.getUserProjects(cu.id);
        const allMembers = await projServ.getProjectsMembers(projectMembers.map(pm => pm.projectId));
        const creatorIds = allMembers.filter(m => m.roleName ==  "requisition_author").map(pm => pm.personId);
        const creators = await ps.findPersonsByIds(creatorIds);
        const drafts = await ds.getDrafts();
        const dm : Map<string, RequisitionDraft> = new Map();
        drafts.filter(d => d.emailId && d.emailId.trim().length > 0).forEach(draft => dm.set(draft.emailId, draft));
        setDraftMap(dm);
        emails = creators.map(p => p.email);
        excelLog("emails = " + emails.join(" , "));

        await excelLog("Before loading email");
        const data = await MSGraphService.getInstance().getMessages(emails);
        await excelLog("After loading email " + JSON.stringify(data as any));
        setEmails(data);
  
        const emailIds = data.map(email => email.id);
        const requisitions = await rs.findAllByEmailIds(emailIds);
        const rm : Map<string, Requisition> = new Map();
        requisitions.filter(r => r.emailId && r.emailId.trim().length > 0).forEach(r => rm.set(r.emailId, r));
        setRequisitionMap(rm);

  
      });

      /*
      // Получаем access token через Office SSO
      const bootstrapToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });

      // Запрос к Graph API
      const response = await fetch('https://graph.microsoft.com/v1.0/me/messages?$top=5&$orderby=receivedDateTime desc', {
        headers: {
          Authorization: `Bearer ${bootstrapToken}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${await response.text()}`);
      }

      const data = await response.json();
      */

    } catch (err: any) {
      await excelLog("Error " + err);
      console.error('SSO error:', err);
      setError(`Ошибка: ${err.message || 'Не удалось получить доступ к Outlook. Проверьте настройки в Azure.'}`);
    } finally {
      setLoading(false);
    }
  };

  // Функция для форматирования даты
  const formatDate = (dateString: string): string => {
    const date = new Date(dateString);
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffDays = Math.floor(diffHours / 24);

    if (diffHours < 1) return 'Только что';
    if (diffHours < 24) return `${diffHours} ч. назад`;
    if (diffDays === 1) return 'Вчера';
    if (diffDays < 7) return `${diffDays} дн. назад`;

    return date.toLocaleDateString('ru-RU', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    });
  };

  // Получение иконки для важности письма
  const getImportanceIcon = (importance: string) => {
    switch (importance) {
      case 'high':
        return <Icon iconName="Important" style={{ color: '#d13438', marginRight: '4px' }} />;
      case 'low':
        return <Icon iconName="Down" style={{ color: '#107c10', marginRight: '4px' }} />;
      default:
        return null;
    }
  };

  // Обрезка текста предпросмотра
  const truncateText = (text: string, maxLength: number = 100): string => {
    return text.length > maxLength ? text.substring(0, maxLength) + '...' : text;
  };

  const openDraft = async (emailId : string) => {
    excelLog("openDraft", emailId);
    const draft = draftMap.get(emailId);
    if(draft.sheetName) {
      await Excel.run(async (context: Excel.RequestContext) => {
        const worksheet = context.workbook.worksheets.getItem(draft.sheetName);
        worksheet.activate();
      });
    }

  }

  const clickMessage = async (email : Message) => {
    await excelLog("Clicked " + email.from.emailAddress.address + "\n" + email.body.content);
    setPrevMessage(message);
    if(message == email) {
      setMessage(undefined);
    } else {
      setMessage(email);
    }
    return 1;
  }

  const createDraft = async (requisition : RequisitionDraft) => {
      await Excel.run(async (context: Excel.RequestContext) => {
        try {
          const ds = await DraftService.create(context);
          await ds.addRequisitionDraft(requisition);
        } catch(e) {
          await excelLog("Error add draft " + e);
        }
      });
  }

  const analyzeMessage = async (email : Message) => {
    let systemMessage = undefined;
    await Excel.run(async (context: Excel.RequestContext) => {
      const ms = await MetadataService.create(context);
      systemMessage = await ms.generateRequisitionSystemMessage();
    });

    if(!systemMessage) {
      throw new Error("Cannot create system message for AI request");
    }
    const response = await AIService.getInstance().callAI(
      `${systemMessage}
       Вот письмо:
       Subject : ${email.subject}
       ${email.body.content}
      `
    );
    const content = JSON.parse(response.choices[0]?.message?.content);
    (content as Requisition).emailId = email.id;
    const s = JSON.stringify(content, null, 2);
    excelLog(s);
    const newAIResults = {...aiResults};
    newAIResults[email.id] = content;
    setAIResults(newAIResults);
    // (email as any).aiResult = content;
    
  }

  return (
    
    <div className="outlook-emails-container">
      
      <style>{`
        .outlook-emails-container {
          padding: 20px;
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          max-width: 100%;
          overflow-x: hidden;
        }
        .email-card {
          margin-bottom: 12px;
          border: 1px solid #e1e5e9;
          transition: all 0.2s ease;
          cursor: pointer;
        }
        .email-card:hover {
          // border-color: #0078d4;
          //background-color: #DDDDDD;
          // box-shadow: 0 2px 8px rgba(0, 120, 212, 0.1);
        }
        .email-card:hover .email-header {
          background-color: #DDDDDD;
        }
        .email-card.unread {
          border-left: 4px solid #0078d4;
          background-color: #f8f9fa;
        }
        .email-header {
          // display: flex;
          // justify-content: space-between;
          // align-items: center;
          // margin-bottom: 8px;
        }
        .email-subject {
          font-weight: 600;
          color: #323130;
          margin: 0;
          font-size: 14px;
        }
        .email-from {
          color: #605e5c;
          font-size: 13px;
          margin: 0;
        }
        .email-date {
          color: #8a8886;
          font-size: 12px;
        }
        .email-preview {
          color: #605e5c;
          font-size: 12px;
          margin-top: 8px;
          line-height: 1.3;
        }
        .email-badges {
          display: flex;
          gap: 4px;
          margin-top: 8px;
        }
        .refresh-button {
          margin-bottom: 16px;
        }
        .body {
          max-height: 0px;
          overflow:auto;
          transition: max-height 0.3s ease-out;
        }
        div.email-preview.open {
          height: 0px;
          background-color:yellow;
          opacity:0;
          transition: height 0.3s ease-out;
          transition: opacity 0.3s ease-out;
        }
        div.body.open {
           max-height: 600px; /* Установите достаточно большую высоту для контента */
        }
      `}</style>

      {(!message || true) && (
      <div>
      <div className="d-flex justify-content-between align-items-center mb-3">
        <h5 className="mb-0">
          <Icon iconName="Mail" style={{ marginRight: '8px', color: '#0078d4' }} />
          Последние письма
        </h5>
        <Button
          variant="outline-primary"
          size="sm"
          onClick={fetchEmails}
          disabled={loading}
          className="refresh-button"
        >
          
          {loading ? (
            <>
              <Spinner animation="border" size="sm" className="me-2" />
              Загрузка...
            </>
          ) : (
            <>
              <Icon iconName="Refresh" style={{ marginRight: '4px' }} />
              Обновить
            </>
          )}
        </Button>
      </div>

      {error && (
        <Alert variant="danger" className="mb-3">
          <Icon iconName="ErrorBadge" style={{ marginRight: '8px' }} />
          {error}
        </Alert>
      )}

      {loading ? (
        <div className="text-center py-4">
          <Spinner animation="border" className="mb-3" />
          <p className="text-muted">Загрузка писем...</p>
        </div>
      ) : emails.length > 0 ? (
        <div>
          {emails.map((email) => (
            <Card key={email.id} className={`email-card ${!email.isRead ? 'unread-nouse' : ''}`} >
              <Card.Body className="p-3">
                <div className="email-header" onClick={() => clickMessage(email)}>
                  <div className="flex-grow-1">
                    <h6 className="email-subject">
                      {getImportanceIcon(email.importance)}
                      {email.subject || '(Без темы)'}
                      {!email.isRead && false && (
                        <Badge bg="primary" className="ms-2" style={{ fontSize: '10px' }}>
                          Новое
                        </Badge>
                      )}

                      {draftMap.get(email.id) && 
                        (<Button
                          variant="outline-primary"
                          size="sm"
                          onClick={() => openDraft(email.id)}
                          disabled={loading}
                          className="refresh-button"
                        >Draft is done
                        </Button>)}
                      {requisitionMap.get(email.id) && 
                        (<Button
                          variant="outline-primary"
                          size="sm"
                          onClick={() => openDraft(email.id)}
                          disabled={loading}
                          className="refresh-button"
                        >Requisition is created
                        </Button>)}

                    </h6>
                    <p className="email-from">
                      От: {email.from?.emailAddress?.name || email.from?.emailAddress?.address || 'Неизвестный отправитель'}
                    </p>
                  </div>
                  <div className="email-date">{formatDate(email.receivedDateTime)}</div>
                </div>

                {email.bodyPreview && <div className={`email-preview ${email == message ? 'open' : ''}`} onClick={() => clickMessage(email)}>{truncateText(email.bodyPreview)}</div>}

                <div className="email-badges">
                  {email.hasAttachments && (
                    <Badge bg="secondary" style={{ fontSize: '10px' }}>
                      <Icon iconName="Attach" style={{ marginRight: '2px' }} />
                      Вложения
                    </Badge>
                  )}
                </div>
                  <div className={`body ${email == message ? 'open' : ''}`}>
                    <div>
                    <PrimaryButton
                      text="Analyze"
                      onClick={() => analyzeMessage(email)}
                      iconProps={{ iconName: 'MailTentative' }}
                      />                              

                    </div>
                    {(email == message || email == prevMessage) && (

                      <EmailMessageViewer message={email}></EmailMessageViewer> 
                    )}

                  </div>

                  {(aiResults[email.id]) && (
                    <div>
                      <RequisitionView data={(aiResults[email.id] as Requisition)}></RequisitionView>
                      <div>
                        <PrimaryButton
                            text="Create Requisition Draft"
                            aria-label="Create Requisition Draft to"
                            onClick={() => createDraft(aiResults[email.id] as RequisitionDraft)}
                            iconProps={{ iconName: 'AddToShoppingList' }}
                            />          
                      </div>                      
                    </div>
                  )}

              </Card.Body>
            </Card>
          ))}
        </div>
      ) : (
        <div className="text-center py-4">
          <Icon iconName="Mail" style={{ fontSize: '48px', color: '#8a8886' }} />
          <p className="text-muted mt-3">Письма не найдены</p>
          <Button variant="outline-primary" onClick={fetchEmails}>
            Попробовать снова
          </Button>
        </div>
      )}
      </div>
    )} 

    </div>

  );
};

export default OutlookEmails;