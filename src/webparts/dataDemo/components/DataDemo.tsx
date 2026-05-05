// ABOUTME: Main component for the DataDemo web part displaying a CRUD table.
// ABOUTME: Two-tier Pivot tabs select transport (REST/PnPjs) and endpoint (SharePoint/Graph/etc).

import * as React from 'react';
import styles from './DataDemo.module.scss';
import type { IDataDemoProps } from './IDataDemoProps';
import { IEventItem, ISpeaker, SessionType } from '../models/IEventItem';
import { ISpService } from '../models/ISpService';
import { IJokeService } from '../models/IJokeService';
import { IGraphQueryService } from '../models/IGraphQueryService';
import { Transport, Endpoint } from '../services/ServiceFactory';
import JokePanel from './JokePanel';
import GraphExplorer from './GraphExplorer';
import { Logger } from '../utilities/logger';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  PrimaryButton,
  DefaultButton,
  TextField,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Stack,
  IStackTokens,
  IconButton,
  Dialog,
  DialogType,
  DialogFooter,
  Pivot,
  PivotItem,
  DatePicker,
  Dropdown,
  IDropdownOption
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';

const SESSION_TYPE_OPTIONS: IDropdownOption[] = [
  { key: '30 minute session', text: '30 minute session' },
  { key: '45 minute session', text: '45 minute session' },
  { key: '50 minute session', text: '50 minute session' },
  { key: '60 minute session', text: '60 minute session' },
  { key: '70 minute session', text: '70 minute session' },
  { key: '75 minute session', text: '75 minute session' },
  { key: 'Half day workshop', text: 'Half day workshop' },
  { key: 'Full day workshop', text: 'Full day workshop' }
];

const stackTokens: IStackTokens = { childrenGap: 10 };

const PLACEHOLDER_ENDPOINTS: Endpoint[] = ['Simple Auth', 'Entra App'];

const QUERY_PARAM_TRANSPORT = 'transport';
const QUERY_PARAM_ENDPOINT = 'endpoint';

const TRANSPORT_SLUG: Record<Transport, string> = {
  'SPFx': 'spfx',
  'PnPjs': 'pnpjs'
};
const ENDPOINT_SLUG: Record<Endpoint, string> = {
  'SharePoint': 'sharepoint',
  'MS Graph (SP)': 'graphsp',
  'Anonymous': 'anonymous',
  'MS Graph (Explorer)': 'graphexplorer',
  'Simple Auth': 'simpleauth',
  'Entra App': 'entraapp'
};

const slugTo = <T extends string>(map: Record<T, string>, slug: string | undefined): T | undefined =>
  slug ? (Object.keys(map) as T[]).find((k) => map[k] === slug) : undefined;

const readDemoQueryParams = (): { transport?: Transport; endpoint?: Endpoint } => {
  const params = new URLSearchParams(window.location.search);
  return {
    transport: slugTo(TRANSPORT_SLUG, params.get(QUERY_PARAM_TRANSPORT) ?? undefined),
    endpoint: slugTo(ENDPOINT_SLUG, params.get(QUERY_PARAM_ENDPOINT) ?? undefined)
  };
};

const DataDemo: React.FC<IDataDemoProps> = ({ factory, site, list }) => {
  const [items, setItems] = React.useState<IEventItem[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [showDialog, setShowDialog] = React.useState(false);
  const [editItem, setEditItem] = React.useState<IEventItem>({ Title: '' });
  const [isEditing, setIsEditing] = React.useState(false);
  const [transport, setTransport] = React.useState<Transport>(() => readDemoQueryParams().transport ?? 'SPFx');
  const [endpoint, setEndpoint] = React.useState<Endpoint>(() => {
    const q = readDemoQueryParams();
    const t = q.transport ?? 'SPFx';
    const e = q.endpoint ?? 'SharePoint';
    // The free-form Graph endpoint is SPFx-only; fall back if the URL combo is invalid.
    return e === 'MS Graph (Explorer)' && t !== 'SPFx' ? 'SharePoint' : e;
  });
  const [spService, setSpService] = React.useState<ISpService | undefined>(undefined);
  const [jokeService, setJokeService] = React.useState<IJokeService | undefined>(undefined);
  const [graphQueryService, setGraphQueryService] = React.useState<IGraphQueryService | undefined>(undefined);

  const isAnonymous = endpoint === 'Anonymous';

  const loadItems = React.useCallback(async (svc: ISpService, lst: typeof list): Promise<void> => {
    Logger.info(`loadItems: requesting items from list ${lst?.title ?? '(none)'}`);
    setLoading(true);
    setError(undefined);

    try {
      const result = await svc.getItems(lst ?? { title: '', id: '' });
      Logger.debug(`loadItems: received ${result.length} item(s)`, result);
      setItems(result);
      setLoading(false);
    } catch (err) {
      Logger.error(`loadItems failed: ${(err as Error).message}`, err);
      setLoading(false);
      setError(`Failed to load items: ${(err as Error).message}`);
    }
  }, []);

  const initServiceAndLoad = React.useCallback(async (
    t: Transport, ep: Endpoint, f: typeof factory, s: typeof site, l: typeof list
  ): Promise<void> => {
    const anon = ep === 'Anonymous';
    const graphQuery = ep === 'MS Graph (Explorer)';
    const needsSite = !anon && !graphQuery;

    Logger.debug(`initServiceAndLoad: transport=${t}, endpoint=${ep}`);

    if (!f || (needsSite && (!s || !l))) {
      Logger.debug('initServiceAndLoad: skipping (factory/site/list missing)');
      setSpService(undefined);
      setJokeService(undefined);
      setGraphQueryService(undefined);
      setItems([]);
      return;
    }

    setLoading(true);
    setError(undefined);

    try {
      if (anon) {
        const svc = f.createJokeService(t);
        setSpService(undefined);
        setGraphQueryService(undefined);
        setJokeService(svc);
        setLoading(false);
      } else if (graphQuery) {
        const svc = await f.createGraphQueryService();
        setSpService(undefined);
        setJokeService(undefined);
        setGraphQueryService(svc);
        setLoading(false);
      } else {
        const svc = await f.createSpService(t, ep, s ?? { url: '', id: '' });
        setJokeService(undefined);
        setGraphQueryService(undefined);
        setSpService(svc);
        setLoading(false);
        await loadItems(svc, l);
      }
    } catch (err) {
      Logger.error(`initServiceAndLoad failed: ${(err as Error).message}`, err);
      setLoading(false);
      setError(`Failed to initialize service: ${(err as Error).message}`);
    }
  }, [loadItems]);

  const onTransportChanged = React.useCallback((item?: PivotItem): void => {
    if (!item) return;
    const newTransport = item.props.itemKey as Transport;
    setTransport((prev) => {
      if (prev === newTransport) return prev;
      Logger.info(`transport changed: ${prev} -> ${newTransport}`);
      setSpService(undefined);
      setJokeService(undefined);
      setGraphQueryService(undefined);
      return newTransport;
    });
    // The free-form Graph endpoint is SPFx-only; fall back to SharePoint if hidden.
    if (newTransport === 'PnPjs') {
      setEndpoint((prev) => (prev === 'MS Graph (Explorer)' ? 'SharePoint' : prev));
    }
  }, []);

  const onEndpointChanged = React.useCallback((item?: PivotItem): void => {
    if (!item) return;
    const newEndpoint = item.props.itemKey as Endpoint;
    setEndpoint((prev) => {
      if (prev === newEndpoint) return prev;
      Logger.info(`endpoint changed: ${prev} -> ${newEndpoint}`);
      setSpService(undefined);
      setJokeService(undefined);
      setGraphQueryService(undefined);
      return newEndpoint;
    });
  }, []);

  // Reflect the active transport/endpoint in the URL so a browser refresh restores them.
  React.useEffect(() => {
    const tSlug = TRANSPORT_SLUG[transport];
    const eSlug = ENDPOINT_SLUG[endpoint];
    const params = new URLSearchParams(window.location.search);
    if (params.get(QUERY_PARAM_TRANSPORT) === tSlug && params.get(QUERY_PARAM_ENDPOINT) === eSlug) {
      return;
    }
    params.set(QUERY_PARAM_TRANSPORT, tSlug);
    params.set(QUERY_PARAM_ENDPOINT, eSlug);
    const newUrl = `${window.location.pathname}?${params.toString()}${window.location.hash}`;
    window.history.replaceState(undefined, '', newUrl);
  }, [transport, endpoint]);

  // Initialize on mount and re-initialize when inputs change
  React.useEffect(() => {
    if (PLACEHOLDER_ENDPOINTS.indexOf(endpoint) >= 0) {
      Logger.debug(`endpoint ${endpoint} is a placeholder, skipping service init`);
      return;
    }
    Logger.debug(`init effect: site=${site?.id ?? 'none'}, list=${list?.id ?? 'none'}, transport=${transport}, endpoint=${endpoint}`);
    initServiceAndLoad(transport, endpoint, factory, site, list)
      .catch(() => { /* handled internally */ });
  }, [factory, site?.id, list?.id, transport, endpoint]); // eslint-disable-line react-hooks/exhaustive-deps

  const onAddItem = React.useCallback((): void => {
    setShowDialog(true);
    setEditItem({ Title: '' });
    setIsEditing(false);
  }, []);

  const onEditItem = React.useCallback((item: IEventItem): void => {
    setShowDialog(true);
    setEditItem({ ...item });
    setIsEditing(true);
  }, []);

  const onCloseDialog = React.useCallback((): void => {
    setShowDialog(false);
  }, []);

  const onSaveItem = React.useCallback(async (): Promise<void> => {
    if (!spService || !list) return;

    Logger.info(`onSaveItem: ${isEditing ? 'update' : 'create'} item${isEditing ? ` id=${editItem.Id}` : ''}`);
    setShowDialog(false);
    setLoading(true);

    try {
      if (isEditing && editItem.Id) {
        const updated = await spService.updateItem(list, editItem.Id, editItem);
        Logger.debug('onSaveItem update result:', updated);
      } else {
        const created = await spService.createItem(list, editItem);
        Logger.debug('onSaveItem create result:', created);
      }
      await loadItems(spService, list);
    } catch (err) {
      Logger.error(`onSaveItem failed: ${(err as Error).message}`, err);
      setLoading(false);
      setError(`Failed to save item: ${(err as Error).message}`);
    }
  }, [spService, list, isEditing, editItem, loadItems]);

  const onDeleteItem = React.useCallback(async (id: number): Promise<void> => {
    if (!spService || !list) return;

    Logger.info(`onDeleteItem: deleting id=${id}`);
    setLoading(true);

    try {
      await spService.deleteItem(list, id);
      Logger.debug(`onDeleteItem: deleted id=${id}`);
      await loadItems(spService, list);
    } catch (err) {
      Logger.error(`onDeleteItem failed: ${(err as Error).message}`, err);
      setLoading(false);
      setError(`Failed to delete item: ${(err as Error).message}`);
    }
  }, [spService, list, loadItems]);

  const renderPlaceholder = (): React.ReactElement => {
    return (
      <div className={styles.placeholder} data-automation-id={`dataDemo-placeholder-${endpoint}`}>
        <Stack tokens={stackTokens} horizontalAlign="center">
          <h3>{endpoint}</h3>
          <p>This demo is not yet implemented. Check back soon.</p>
        </Stack>
      </div>
    );
  };

  const renderCrudPanel = (): React.ReactElement => {
    if (!spService || !list) {
      return (
        <Spinner size={SpinnerSize.large} label="Initializing..." data-automation-id="dataDemo-spinner-init" />
      );
    }

    const columns: IColumn[] = [
      {
        key: 'actions',
        name: 'Actions',
        minWidth: 60,
        maxWidth: 70,
        onRender: (item: IEventItem) => (
          <Stack horizontal tokens={{ childrenGap: 4 }}>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit"
              ariaLabel="Edit item"
              data-automation-id={`dataDemo-button-edit-${item.Id}`}
              onClick={() => onEditItem(item)}
            />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete"
              ariaLabel="Delete item"
              data-automation-id={`dataDemo-button-delete-${item.Id}`}
              onClick={() => onDeleteItem(item.Id!)}
            />
          </Stack>
        )
      },
      {
        key: 'SessionDate',
        name: 'Date',
        fieldName: 'SessionDate',
        minWidth: 70,
        maxWidth: 80,
        isResizable: true,
        onRender: (item: IEventItem) =>
          item.SessionDate ? new Date(item.SessionDate).toLocaleDateString() : ''
      },
      { key: 'Title', name: 'Event', fieldName: 'Title', minWidth: 180, maxWidth: 360, isResizable: true, isMultiline: true },
      { key: 'Session', name: 'Session', fieldName: 'Session', minWidth: 200, maxWidth: 400, isResizable: true, isMultiline: true },
      { key: 'SessionType', name: 'Type', fieldName: 'SessionType', minWidth: 90, maxWidth: 130, isResizable: true, isMultiline: true },
      {
        key: 'Speaker',
        name: 'Speaker',
        minWidth: 100,
        maxWidth: 160,
        isResizable: true,
        isMultiline: true,
        onRender: (item: IEventItem) =>
          item.Speaker?.map((s) => s.Title).join(', ') ?? ''
      },
      {
        key: 'EventSite',
        name: 'Event Site',
        minWidth: 140,
        maxWidth: 240,
        isResizable: true,
        onRender: (item: IEventItem) =>
          item.EventSite?.Url ? (
            <a href={item.EventSite.Url} target="_blank" rel="noreferrer">
              {item.EventSite.Description || item.EventSite.Url}
            </a>
          ) : null
      }
    ];

    return (
      <>
        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => setError(undefined)}
            data-automation-id="dataDemo-message-error"
          >
            {error}
          </MessageBar>
        )}

        <h2 className={styles.viewHeading} data-automation-id="dataDemo-heading-view">Upcoming Events</h2>

        <Stack horizontal tokens={stackTokens}>
          <PrimaryButton
            text="Add Item"
            iconProps={{ iconName: 'Add' }}
            onClick={onAddItem}
            data-automation-id="dataDemo-button-add"
          />
          <DefaultButton
            text="Refresh"
            iconProps={{ iconName: 'Refresh' }}
            onClick={() => loadItems(spService, list)}
            data-automation-id="dataDemo-button-refresh"
          />
        </Stack>

        {loading ? (
          <Spinner size={SpinnerSize.large} label="Loading items..." data-automation-id="dataDemo-spinner-loading" />
        ) : (
          <DetailsList
            items={items}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            data-automation-id="dataDemo-list-items"
          />
        )}

        <Dialog
          hidden={!showDialog}
          onDismiss={onCloseDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: isEditing ? 'Edit Event' : 'Add Event'
          }}
          modalProps={{ isBlocking: true }}
          minWidth={520}
        >
          <Stack tokens={stackTokens} key={isEditing ? `edit-${editItem.Id ?? 'new'}` : 'add'}>
            <TextField
              label="Event"
              value={editItem.Title}
              onChange={(_e, val) => setEditItem({ ...editItem, Title: val || '' })}
              required
              data-automation-id="dataDemo-input-title"
            />
            <TextField
              label="Session"
              value={editItem.Session ?? ''}
              onChange={(_e, val) => setEditItem({ ...editItem, Session: val || '' })}
              data-automation-id="dataDemo-input-session"
            />
            <DatePicker
              label="Date"
              textField={{
                onRenderLabel: () => (
                  <label className={styles.required}>
                    Date<span aria-hidden="true"> *</span>
                  </label>
                )
              }}
              value={editItem.SessionDate ? new Date(editItem.SessionDate) : undefined}
              onSelectDate={(d) =>
                setEditItem({ ...editItem, SessionDate: d ? d.toISOString() : undefined })
              }
              data-automation-id="dataDemo-input-sessionDate"
            />
            <Dropdown
              label="Session Type"
              selectedKey={editItem.SessionType}
              options={SESSION_TYPE_OPTIONS}
              onChange={(_e, opt) =>
                setEditItem({ ...editItem, SessionType: opt?.key as SessionType })
              }
              data-automation-id="dataDemo-input-sessionType"
            />
            {factory && (
              <PeoplePicker
                context={{
                  absoluteUrl: factory.context.pageContext.web.absoluteUrl,
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  msGraphClientFactory: factory.context.msGraphClientFactory as any,
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  spHttpClient: factory.context.spHttpClient as any
                }}
                webAbsoluteUrl={site?.url}
                titleText={endpoint === 'MS Graph (SP)' ? 'Speaker (read-only on Graph)' : 'Speaker'}
                personSelectionLimit={10}
                principalTypes={[PrincipalType.User]}
                resolveDelay={300}
                disabled={endpoint === 'MS Graph (SP)'}
                defaultSelectedUsers={editItem.Speaker?.map((s) => s.EMail ?? s.Title) ?? []}
                onChange={(items) => {
                  const speakers: ISpeaker[] = items.map((p) => ({
                    Id: parseInt(p.id ?? '0', 10),
                    Title: p.text ?? '',
                    EMail: p.secondaryText
                  }));
                  setEditItem({ ...editItem, Speaker: speakers });
                }}
                ensureUser
                data-automation-id="dataDemo-input-speaker"
              />
            )}
            <TextField
              label="Event Site URL"
              value={editItem.EventSite?.Url ?? ''}
              onChange={(_e, val) =>
                setEditItem({
                  ...editItem,
                  EventSite: { Url: val || '', Description: val || '' }
                })
              }
              data-automation-id="dataDemo-input-eventSiteUrl"
            />
          </Stack>
          <DialogFooter>
            <PrimaryButton
              text="Save"
              onClick={onSaveItem}
              disabled={!editItem.Title || !editItem.SessionDate}
              data-automation-id="dataDemo-button-save"
            />
            <DefaultButton
              text="Cancel"
              onClick={onCloseDialog}
              data-automation-id="dataDemo-button-cancel"
            />
          </DialogFooter>
        </Dialog>
      </>
    );
  };

  return (
    <div className={styles.dataDemo} data-automation-id="dataDemo-container-root">
      <Stack tokens={stackTokens}>
        <div className={styles.pivotWrapper}>
          <Pivot
            selectedKey={transport}
            onLinkClick={onTransportChanged}
            data-automation-id="dataDemo-pivot-transport"
          >
            <PivotItem headerText="SPFx" itemKey="SPFx" />
            <PivotItem headerText="PnPjs" itemKey="PnPjs" />
          </Pivot>
        </div>

        <div className={styles.pivotWrapper}>
          <Pivot
            selectedKey={endpoint}
            onLinkClick={onEndpointChanged}
            data-automation-id="dataDemo-pivot-endpoint"
          >
            <PivotItem headerText="Anonymous" itemKey="Anonymous" />
            <PivotItem headerText="SharePoint" itemKey="SharePoint" />
            <PivotItem headerText="MS Graph (SP)" itemKey="MS Graph (SP)" />
            {transport === 'SPFx' && <PivotItem headerText="MS Graph" itemKey="MS Graph (Explorer)" />}
            <PivotItem headerText="Simple Auth" itemKey="Simple Auth" />
            <PivotItem headerText="Entra App" itemKey="Entra App" />
          </Pivot>
        </div>

        {PLACEHOLDER_ENDPOINTS.indexOf(endpoint) >= 0
          ? renderPlaceholder()
          : endpoint === 'MS Graph (Explorer)' && graphQueryService
            ? <GraphExplorer service={graphQueryService} />
            : isAnonymous && jokeService
              ? <JokePanel service={jokeService} />
              : renderCrudPanel()
        }
      </Stack>
    </div>
  );
};

export default DataDemo;
