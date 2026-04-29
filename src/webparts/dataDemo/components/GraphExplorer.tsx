// ABOUTME: Free-form Graph endpoint explorer for the DataDemo web part.
// ABOUTME: Lets the user pick or type a v1.0 path, run it, and view the JSON response.

import * as React from 'react';
import {
  Stack,
  IStackTokens,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import { IGraphQueryService } from '../models/IGraphQueryService';
import styles from './DataDemo.module.scss';

const stackTokens: IStackTokens = { childrenGap: 10 };

const PRESET_PATHS: IDropdownOption[] = [
  { key: '/me', text: '/me' },
  { key: '/me/messages', text: '/me/messages' },
  { key: '/me/drive/root/children', text: '/me/drive/root/children' },
  { key: '/sites/root', text: '/sites/root' },
  { key: '/users', text: '/users' }
];

const DEFAULT_PATH = '/me';

export interface IGraphExplorerProps {
  service: IGraphQueryService;
}

const GraphExplorer: React.FC<IGraphExplorerProps> = ({ service }) => {
  const [path, setPath] = React.useState<string>(DEFAULT_PATH);
  const [loading, setLoading] = React.useState(false);
  const [response, setResponse] = React.useState<unknown>(undefined);
  const [errorBody, setErrorBody] = React.useState<unknown>(undefined);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  const onPresetChanged = React.useCallback((_e: unknown, option?: IDropdownOption): void => {
    if (option) {
      setPath(option.key as string);
    }
  }, []);

  const onPathChanged = React.useCallback((_e: unknown, val?: string): void => {
    setPath(val ?? '');
  }, []);

  const onRun = React.useCallback(async (): Promise<void> => {
    if (!path) return;

    setLoading(true);
    setResponse(undefined);
    setErrorBody(undefined);
    setErrorMessage(undefined);

    try {
      const result = await service.runQuery(path);
      setResponse(result);
    } catch (err) {
      const e = err as { message?: string; statusCode?: number; body?: unknown; code?: string };
      setErrorMessage(e.message ?? 'Request failed');
      // Graph client errors typically expose body as a JSON string
      let parsedBody: unknown = e.body;
      if (typeof parsedBody === 'string') {
        try {
          parsedBody = JSON.parse(parsedBody);
        } catch {
          // keep as string
        }
      }
      setErrorBody(parsedBody ?? { code: e.code, message: e.message, statusCode: e.statusCode });
    } finally {
      setLoading(false);
    }
  }, [service, path]);

  const onPathKeyDown = React.useCallback((e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter' && !loading && path) {
      e.preventDefault();
      onRun().catch(() => undefined);
    }
  }, [loading, path, onRun]);

  return (
    <Stack tokens={stackTokens} data-automation-id="dataDemo-container-graphExplorer">
      <Stack horizontal tokens={stackTokens} verticalAlign="end">
        <Dropdown
          label="Preset"
          selectedKey={PRESET_PATHS.some((p) => p.key === path) ? path : undefined}
          options={PRESET_PATHS}
          onChange={onPresetChanged}
          styles={{ root: { minWidth: 220 } }}
          data-automation-id="dataDemo-dropdown-graphPreset"
        />
        <TextField
          label="Graph v1.0 path"
          value={path}
          onChange={onPathChanged}
          onKeyDown={onPathKeyDown}
          styles={{ root: { flexGrow: 1 } }}
          data-automation-id="dataDemo-input-graphPath"
        />
        <PrimaryButton
          text="Run"
          iconProps={{ iconName: 'Play' }}
          onClick={onRun}
          disabled={loading || !path}
          data-automation-id="dataDemo-button-graphRun"
        />
      </Stack>

      {errorMessage && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setErrorMessage(undefined)}
          data-automation-id="dataDemo-message-graphError"
        >
          {errorMessage}
        </MessageBar>
      )}

      {loading && (
        <Spinner size={SpinnerSize.large} label="Calling Graph..." data-automation-id="dataDemo-spinner-graph" />
      )}

      {!loading && errorBody !== undefined && (
        <pre className={styles.graphErrorJson} data-automation-id="dataDemo-pre-graphError">
          {JSON.stringify(errorBody, null, 2)}
        </pre>
      )}

      {!loading && response !== undefined && (
        <pre className={styles.graphJson} data-automation-id="dataDemo-pre-graphResponse">
          {JSON.stringify(response, null, 2)}
        </pre>
      )}
    </Stack>
  );
};

export default GraphExplorer;
