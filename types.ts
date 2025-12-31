
export interface RosterRow {
  [key: string]: any;
}

export interface RosterSheet {
  name: string;
  data: any[][];
}

export interface StaffMapping {
  original: string;
  updated: string;
}

export interface ProcessedData {
  sheets: RosterSheet[];
  majorityMonth: string;
}
