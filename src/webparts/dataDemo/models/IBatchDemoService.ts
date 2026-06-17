// ABOUTME: Capability contract for the PnPjs $batch lifecycle demo.
// ABOUTME: Implemented by the PnPjs services only; the UI feature-detects runBatchDemo.

import { IListIdentifier } from './ISpService';

// One phase of the demo (cleanup / create / update / delete-odd). `requests` is
// the number of $batch HTTP calls — operations split by the chosen batch size.
export interface IBatchDemoPhase {
  name: string;
  label: string;
  operations: number;
  requests: number;
}

export interface IBatchDemoResult {
  phases: IBatchDemoPhase[];
  totalOperations: number;
  totalRequests: number;
  batchSize: number;
}

// A scripted create/update/delete lifecycle, each phase sent as one $batch, to
// show the real payoff of batching: many operations in a handful of round-trips.
export interface IBatchDemoService {
  runBatchDemo(
    list: IListIdentifier,
    batchSize: number,
    onPhase?: (phase: IBatchDemoPhase) => Promise<void> | void
  ): Promise<IBatchDemoResult>;
}
