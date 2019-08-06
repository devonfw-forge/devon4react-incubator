import { ProjectData } from './ProjectData';

export interface Employee {
  name: string;
  worksheetData: ProjectData[];
  cell: string;
  total: number;
}
