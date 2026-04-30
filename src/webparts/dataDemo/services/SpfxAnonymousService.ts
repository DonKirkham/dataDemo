// ABOUTME: Anonymous external API calls using SPFx HttpClient (no auth headers).
// ABOUTME: Fetches jokes from a public API to demonstrate unauthenticated HTTP requests.

import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { IJoke, IJokeService } from '../models/IJokeService';

interface IJokeResponse {
  id: number;
  type: string;
  setup: string;
  punchline: string;
}

const JOKE_URL = 'https://official-joke-api.appspot.com/random_joke';

export class SpfxAnonymousService implements IJokeService {
  public readonly url = JOKE_URL;

  constructor(private httpClient: HttpClient) {}

  public async getJoke(): Promise<IJoke> {
    const response: HttpClientResponse = await this.httpClient.get(
      this.url,
      HttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch joke: ${response.statusText}`);
    }

    const joke: IJokeResponse = await response.json();
    return { id: joke.id, setup: joke.setup, punchline: joke.punchline };
  }
}
