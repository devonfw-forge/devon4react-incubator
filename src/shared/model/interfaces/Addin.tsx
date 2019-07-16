import { IEmpleado } from './Empleado';
import { IFieldDef } from './FieldDef';

export interface Addin {
  empleado: IEmpleado;
  fields: IFieldDef[];
}
