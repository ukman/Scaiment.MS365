// RequisitionView.tsx
import * as React from "react";
import {
  Stack,
  Text,
  Label,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  TooltipHost,
  Icon,
  Separator,
  useTheme,
  MessageBar,
  MessageBarType,
  mergeStyleSets,
} from "@fluentui/react";
import { Requisition, RequisitionItem } from "../../util/data/DBSchema";

type RequisitionItem2 = {
  productName: string;
  quantity: number;
  unit: string;
  category: string;
  commentInternal?: string;
  comment?: string;
};

export type Requisition2 = {
  name: string;
  createdBy: string;
  dueDate?: string; // ISO
  commentInternal?: string;
  comment?: string;
  requisitionItems: RequisitionItem2[];
};

export type RequisitionViewProps = {
  data: Requisition;
  className?: string;
};

const classes = mergeStyleSets({
  root: {
    width: "100%",
    maxWidth: 820, // не фикс — просто ограничитель, убери при желании
    margin: "0 auto",
  },
  headerRow: {
    rowGap: 6,
  },
  meta: {
    columnGap: 12,
    rowGap: 6,
    flexWrap: "wrap",
  },
  metaItem: {
    minWidth: 140,
  },
  pill: {
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
    borderRadius: 999,
    padding: "2px 8px",
    border: "1px solid",
    fontSize: 12,
    lineHeight: "18px",
  },
  list: {
    // обеспечивает горизонтальный скролл, если слишком узко
    overflowX: "auto",
  },
  qtyCell: {
    whiteSpace: "nowrap",
  },
  muted: {
    opacity: 0.8,
  },
});

const formatDate = (iso?: Date) => {
    return "" + iso;
}

const formatDate2 = (iso?: string) => {
  if (!iso) return "—";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso; // не валидное значение — отобразим как есть
  return new Intl.DateTimeFormat("ru-RU", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(d);
};

export const RequisitionView: React.FC<RequisitionViewProps> = ({
  data,
  className,
}) => {
  const theme = useTheme();

  const columns = React.useMemo<IColumn[]>(
    () => [
      {
        key: "product",
        name: "Материал",
        fieldName: "productName",
        minWidth: 180,
        maxWidth: 360,
        isMultiline: true,
        isResizable: true,
        onRender: (item: RequisitionItem) => (
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="mediumPlus">{item.productName}</Text>
            <Text variant="small" className={classes.muted}>
              Категория: {item.category || "—"}
            </Text>
            {(item.comment || item.commentInternal) && (
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                {item.comment && (
                  <TooltipHost content="Комментарий для поставщика">
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                      <Icon iconName="Comment" />
                      <Text variant="small">{item.comment}</Text>
                    </Stack>
                  </TooltipHost>
                )}
                {item.commentInternal && (
                  <TooltipHost content="Внутренний комментарий">
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                      <Icon iconName="Lock" />
                      <Text variant="small">{item.commentInternal}</Text>
                    </Stack>
                  </TooltipHost>
                )}
              </Stack>
            )}
          </Stack>
        ),
      },
      {
        key: "qty",
        name: "Кол-во",
        fieldName: "quantity",
        minWidth: 90,
        maxWidth: 110,
        isResizable: true,
        onRender: (item: RequisitionItem) => (
          <Text className={classes.qtyCell}>
            {item.quantity} {item.unit}
          </Text>
        ),
      },
    ],
    []
  );

  const hasAnyExternalComment = false;
    // !!data.comment && data.comment.trim().length > 0;
  const hasAnyInternalComment = false;
    // !!data.commentInternal && data.commentInternal.trim().length > 0;

  return (
    <Stack className={`${classes.root} ${className || ""}`} tokens={{ childrenGap: 12 }}>
      {/* Заголовок */}
      <Stack tokens={{ childrenGap: 6 }} className={classes.headerRow}>
        <Text variant="xLarge" block styles={{ root: { fontWeight: 600 } }}>
          {data.name || "Заявка"}
        </Text>
        <Stack horizontal wrap className={classes.meta}>
          <Stack horizontal verticalAlign="center" className={classes.metaItem}>
            <div
              className={classes.pill}
              style={{
                borderColor: theme.palette.neutralQuaternaryAlt,
                color: theme.palette.neutralPrimary,
                background: theme.palette.neutralLighterAlt,
              }}
              aria-label="Создал"
              title="Создал"
            >
              <Icon iconName="Contact" />
              <span>{data.createdBy || "—"}</span>
            </div>
          </Stack>

          <Stack horizontal verticalAlign="center" className={classes.metaItem}>
            <div
              className={classes.pill}
              style={{
                borderColor: theme.palette.neutralQuaternaryAlt,
                color: theme.palette.neutralPrimary,
                background: theme.palette.neutralLighterAlt,
              }}
              aria-label="Срок"
              title="Срок"
            >
              <Icon iconName="Calendar" />
              <span>Срок: {formatDate(data.dueDate)}</span>
            </div>
          </Stack>
        </Stack>
      </Stack>

      {/* Комментарии сверху, если есть */}
      {(hasAnyExternalComment || hasAnyInternalComment) && (
        <Stack tokens={{ childrenGap: 8 }}>
          {hasAnyExternalComment && (
            <MessageBar
              messageBarType={MessageBarType.severeWarning}
              isMultiline={true}
            >
              <b>Комментарий к заявке:</b> {(data as any).comment}
            </MessageBar>
          )}
          {hasAnyInternalComment && (
            <MessageBar messageBarType={MessageBarType.info} isMultiline={true}>
              <b>Внутренний комментарий:</b> {(data as any).commentInternal}
            </MessageBar>
          )}
        </Stack>
      )}

      <Separator>Позиции</Separator>

      <div className={classes.list} role="region" aria-label="Список позиций заявки">
        <DetailsList
          items={(data as any).requisitionItems || []}
          columns={columns}
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          compact={true}
          selectionMode={0} // none
          ariaLabelForGrid="Таблица позиций заявки"
        />
      </div>
    </Stack>
  );
};
