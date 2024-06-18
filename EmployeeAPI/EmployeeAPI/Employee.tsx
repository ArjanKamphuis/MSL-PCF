import * as React from "react";

export type EmployeeProps = {
    webApi: ComponentFramework.WebApi;
};

type EmployeeInformation = {
    name: string;
    schedule: string[];
};

const days: string[] = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];

const fetchEmployeeSchedule = (entity?: ComponentFramework.WebApi.Entity): string[] => {
    if (!entity) return [];
    let schedule: string[] = [];
    for (const day of days) {
        if (entity[`bs_${day}begintime`]) {
            const begin: string = entity[`bs_${day}begintime`];
            const end: string = entity[`bs_${day}endtime`];
            schedule.push(`${day[0].toUpperCase()}${day.slice(1)}: ${begin} - ${end}`);
        }
    }
    return schedule;
};

const deskTableName: string = 'bs_desk';
const deskNameColumn: string = 'bs_deskname';
const deskEmployeeColumn: string = 'bs_employee';
const employeeTableName: string = 'bs_employee_mock';
const scheduleTableName: string = 'bs_schedule_mock';
const employeeTableColumnName: string = 'bs_employee';

export const Employee = React.memo((props: EmployeeProps) => {
    const { webApi } = props;

    const [desks, setDesks] = React.useState<ComponentFramework.WebApi.Entity[]>([]);
    const [currentEmployeeId, setCurrentEmployeeId] = React.useState<string>('');
    const [currentEmployee, setCurrentEmployee] = React.useState<EmployeeInformation | null>(null);

    React.useEffect(() => {
        const fetchDesks = async () => {
            try {
                const response: ComponentFramework.WebApi.RetrieveMultipleResponse = await webApi.retrieveMultipleRecords(deskTableName);
                setDesks(response.entities);
            } catch (errorResponse: any) {
                console.error(errorResponse.message);
            }
        };
        fetchDesks();
    }, [webApi]);

    React.useEffect(() => {
        if (!currentEmployeeId) return;
        const fetchEmployee = async () => {
            try {
                const employee: ComponentFramework.WebApi.Entity = await webApi.retrieveRecord(employeeTableName, currentEmployeeId);
                const schedules: ComponentFramework.WebApi.RetrieveMultipleResponse = await webApi.retrieveMultipleRecords(scheduleTableName);
                const schedule: ComponentFramework.WebApi.Entity | undefined = schedules.entities.find(s => s['_bs_employeeid_value'] === employee['bs_employee_mockid']);
                setCurrentEmployee({ name: employee['bs_employee'], schedule: fetchEmployeeSchedule(schedule) });
            } catch (errorResponse: any) {
                console.error(errorResponse.message);
            }
        };
        fetchEmployee();
    }, [currentEmployeeId, webApi]);

    const deskList = React.useMemo((): React.JSX.Element[] => desks.map(desk => {
        return (
            <li key={desk['bs_deskid']}>
                <button onClick={() => setCurrentEmployeeId(desk['_bs_employee_value'])}>{desk['bs_deskname']}</button>
            </li>
        );
    }), [desks]);

    const employeeDetails = React.useMemo((): React.JSX.Element => {
        if (!currentEmployee) return <></>;
        return (
            <div>
                <p>{currentEmployee.name}</p>
                {currentEmployee.schedule.map(day => <li>{day}</li>)}
            </div>
        );
    }, [currentEmployee]);

    return (
        <div>
            <ul>{deskList}</ul>
            {employeeDetails}
        </div>
    );
});