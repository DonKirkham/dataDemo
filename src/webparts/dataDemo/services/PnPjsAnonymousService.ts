// ABOUTME: Anonymous external API calls using @pnp/queryable with BrowserFetch behavior.
// ABOUTME: Demonstrates PnPjs composable pipeline for public endpoints without SharePoint context.

import { Queryable } from '@pnp/queryable';
import { BrowserFetch, JSONParse, ResolveOnData, RejectOnError } from '@pnp/queryable';
import { IJoke, IJokeService } from '../models/IJokeService';

interface IJokeResponse {
  id: number;
  type: string;
  setup: string;
  punchline: string;
}

const JOKE_URL = 'https://official-joke-api.appspot.com/random_joke';

export class PnPjsAnonymousService implements IJokeService {
  public readonly url = JOKE_URL;

  public async getJoke(): Promise<IJoke> {
    const q = new Queryable(this.url);
    q.using(BrowserFetch(), RejectOnError(), ResolveOnData(), JSONParse());
    // Strip PnPjs tracking header to avoid CORS preflight on external APIs
    q.on.pre(async (url, init, result) => {
      const headers = init.headers as Record<string, string> | undefined;
      if (headers) {
        delete headers['X-PnPjs-RequestId'];
      }
      return [url, init, result];
    });

    const joke: IJokeResponse = await q();
    return { id: joke.id, setup: joke.setup, punchline: joke.punchline };
  }
}
