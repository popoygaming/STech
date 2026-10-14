export interface IEmployeeProfile {
  id: number;
  fullName: string;
  role: string;
  birthday?: string;
  employmentStartDate?: string;
  phone?: string;
  email?: string;
  state?: string;
  profileImageUrl?: string;
  availability?: string;
  activity?: string;
}

export default class EmployeeInfo {
  public static readonly listTitle: string = 'Employee Directory';

  public static readonly columns = {
    id: 'Id',
    fullName: 'Title',
    role: 'Role',
    birthday: 'Birthday',
    employmentStartDate: 'EmploymentStartDate',
    phone: 'Phone',
    email: 'Email',
    state: 'State'
  };

  public static readonly selectFields: string[] = [
    EmployeeInfo.columns.id,
    EmployeeInfo.columns.fullName,
    EmployeeInfo.columns.role,
    EmployeeInfo.columns.birthday,
    EmployeeInfo.columns.employmentStartDate,
    EmployeeInfo.columns.phone,
    EmployeeInfo.columns.email,
    EmployeeInfo.columns.state
  ];
}
