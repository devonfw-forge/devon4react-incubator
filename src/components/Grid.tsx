import * as React from 'react';

export const Grid: React.FunctionComponent<any> = (props) => {
  const handleChangeName = (event: any, idx: any) => {
    props.employee.fields[idx].fieldName = event.target.value;

    localStorage.setItem('empleado', JSON.stringify(props.employee));
    props.setEmp(props.employee);
  };
  const handleChangeValue = (event: any, idx: any) => {
    props.employee.fields[idx].fieldValue = event.target.value;

    localStorage.setItem('empleado', JSON.stringify(props.employee));
    props.setEmp(props.employee);
  };

  return props.employee.fields.map((project: any, index: any) => {
    console.log(project.fieldName);
    return (
      <div key={index}>
        <input
          defaultValue={project.fieldName.toString()}
          type="text"
          onChange={(event) => handleChangeName(event, index)}
        />
        <input
          defaultValue={project.fieldValue.toString()}
          type="text"
          onChange={(event) => handleChangeValue(event, index)}
        />
      </div>
    );
  });
};
