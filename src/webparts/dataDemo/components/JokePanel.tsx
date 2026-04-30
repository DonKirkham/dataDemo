// ABOUTME: Displays a single joke with a delayed punchline reveal.
// ABOUTME: Fetches from a public joke API via the provided IJokeService instance.

import * as React from 'react';
import styles from './JokePanel.module.scss';
import { IJokeService } from '../models/IJokeService';
import { Logger } from '../utilities/logger';
import {
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Stack
} from '@fluentui/react';

export interface IJokePanelProps {
  service: IJokeService;
}

const JokePanel: React.FC<IJokePanelProps> = ({ service }) => {
  const [setup, setSetup] = React.useState('');
  const [punchline, setPunchline] = React.useState('');
  const [showPunchline, setShowPunchline] = React.useState(false);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const timerRef = React.useRef<number | undefined>(undefined);

  const loadJoke = React.useCallback((): void => {
    if (timerRef.current) {
      window.clearTimeout(timerRef.current);
      timerRef.current = undefined;
    }

    setLoading(true);
    setError(undefined);
    setShowPunchline(false);

    Logger.info(`loadJoke: GET ${service.url}`);

    service.getJoke().then((joke) => {
      Logger.debug('loadJoke: received joke', joke);
      setSetup(joke.setup);
      setPunchline(joke.punchline);
      setLoading(false);

      timerRef.current = window.setTimeout(() => {
        setShowPunchline(true);
      }, 3000);
    }).catch((err: Error) => {
      Logger.error(`loadJoke failed: ${err.message}`, err);
      setLoading(false);
      setError(`Failed to fetch joke: ${err.message}`);
    });
  }, [service]);

  React.useEffect(() => {
    loadJoke();
    return () => {
      if (timerRef.current) {
        window.clearTimeout(timerRef.current);
      }
    };
  }, [loadJoke]);

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

      <div className={styles.jokePanel} data-automation-id="dataDemo-container-joke">
        {loading ? (
          <Spinner size={SpinnerSize.large} label="Fetching joke..." data-automation-id="dataDemo-spinner-loading" />
        ) : (
          <>
            <div className={styles.setup} data-automation-id="dataDemo-text-setup">{setup}</div>
            {showPunchline && (
              <div className={styles.punchline} data-automation-id="dataDemo-text-punchline">{punchline}</div>
            )}
          </>
        )}
      </div>

      {showPunchline && (
        <Stack horizontalAlign="center" className={styles.nextButton}>
          <DefaultButton
            text="Next Joke"
            iconProps={{ iconName: 'Refresh' }}
            onClick={loadJoke}
            data-automation-id="dataDemo-button-nextjoke"
          />
        </Stack>
      )}
    </>
  );
};

export default JokePanel;
