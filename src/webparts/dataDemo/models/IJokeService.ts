// ABOUTME: Service contract for fetching jokes from a public API.
// ABOUTME: Distinct from ISpService — anonymous endpoints are not SharePoint operations.

export interface IJoke {
  id: number;
  setup: string;
  punchline: string;
}

export interface IJokeService {
  readonly url: string;
  getJoke(): Promise<IJoke>;
}
