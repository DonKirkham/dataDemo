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
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  // Bumped on every successful load so the CSS reveal animations replay per joke.
  const [round, setRound] = React.useState(0);

  const loadJoke = React.useCallback((): void => {
    setLoading(true);
    setError(undefined);

    Logger.info(`loadJoke: GET ${service.url}`);

    service.getJoke().then((joke) => {
      Logger.debug('loadJoke: received joke', joke);
      setSetup(joke.setup);
      setPunchline(joke.punchline);
      setRound((r) => r + 1);
      setLoading(false);
    }).catch((err: Error) => {
      Logger.error(`loadJoke failed: ${err.message}`, err);
      setLoading(false);
      setError(`Failed to fetch joke: ${err.message}`);
    });
  }, [service]);

  React.useEffect(() => {
    loadJoke();
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

      {loading ? (
        <div className={styles.jokePanel} data-automation-id="dataDemo-container-joke">
          <Spinner size={SpinnerSize.large} label="Fetching joke..." data-automation-id="dataDemo-spinner-loading" />
        </div>
      ) : (
        // key={round} remounts the subtree on each joke so the CSS reveal
        // animations (punchline at 2s, button at 3s) restart from the beginning.
        // All timing lives in JokePanel.module.scss — no JS timers.
        <React.Fragment key={round}>
          <div className={styles.jokePanel} data-automation-id="dataDemo-container-joke">
            <div className={styles.setup} data-automation-id="dataDemo-text-setup">{setup}</div>
            <div className={styles.punchline} data-automation-id="dataDemo-text-punchline">{punchline}</div>
          </div>

          <Stack horizontalAlign="center" className={styles.nextButton}>
            <DefaultButton
              text="Next Joke"
              iconProps={{ iconName: 'Refresh' }}
              onClick={loadJoke}
              data-automation-id="dataDemo-button-nextjoke"
            />
          </Stack>
        </React.Fragment>
      )}
    </>
  );
};

export default JokePanel;
