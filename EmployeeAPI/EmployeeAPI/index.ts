import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class EmployeeAPI implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;

    private static _employeeTableName: string = 'bs_employee_mock';
    private static _scheduleTableName: string = 'bs_schedule_mock';
    private static _employeeTableColumnName: string = 'bs_employee';

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this._context = context;
        this._container = document.createElement('div');
        container.appendChild(this._container);

        this._container.appendChild(this.createHTMLButtonElement('Click Me!', () => { this.fetchEmployees()}));
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }

    private createHTMLButtonElement(label: string, onClickHandler: (event?: any) => void): HTMLButtonElement {
        const button = document.createElement('button');
        button.textContent = label;
        button.onclick = onClickHandler;
        return button;
    }

    private async fetchEmployees(): Promise<void> {
        try {
            const response: ComponentFramework.WebApi.RetrieveMultipleResponse = await this._context.webAPI.retrieveMultipleRecords(EmployeeAPI._employeeTableName);
            for (const employee of response.entities) {
                const div = document.createElement('div');
                div.innerText = employee[EmployeeAPI._employeeTableColumnName];
                this._container.appendChild(div);
            }
        } catch (errorResponse: any) {
            console.error(errorResponse.message);
        }
    }
}
