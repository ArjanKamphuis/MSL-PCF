import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class TSWebAPI implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _controlViewRendered: boolean;

    private _createEntity1Button: HTMLButtonElement;
    private _createEntity2Button: HTMLButtonElement;
    private _createEntity3Button: HTMLButtonElement;
    private _deleteRecordButton: HTMLButtonElement;
    private _fetchXmlRefreshButton: HTMLButtonElement;
    private _oDataRefreshButton: HTMLButtonElement;

    private _oDataStatusContainerDiv: HTMLDivElement;
    private _resultContainerDiv: HTMLDivElement;

    private static _entityName: string = 'account';
    private static _requiredAttributeName: string = 'name';
    private static _requiredAttributeValue: string = 'Web API Custom Control (Sample)';
    private static _currencyAttributeName: string = 'revenue';
    private static _currencyAttributeNameFriendlyName: string = 'annual revenue';

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
        this._controlViewRendered = false;
        this._container = document.createElement('div');
        this._container.classList.add('TSWebAPI_Container');
        container.appendChild(this._container);
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        if (!this._controlViewRendered) {
            this._controlViewRendered = true;

            this.renderCreateExample();
            this.renderDeleteExample();
            this.renderFetchXmlRetrieveMultipleExample();
            this.renderODataRetrieveMultipleExample();

            this.renderResultsDiv();
        }
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
        this._createEntity1Button.removeEventListener('click', this.createButtonOnClickHandler);
        this._createEntity2Button.removeEventListener('click', this.createButtonOnClickHandler);
        this._createEntity3Button.removeEventListener('click', this.createButtonOnClickHandler);
        this._deleteRecordButton.removeEventListener('click', this.deleteButtonOnClickHandler);
        this._fetchXmlRefreshButton.removeEventListener('click', this.calculateAverageButtonOnClickHandler);
        this._oDataRefreshButton.removeEventListener('click', this.refreshRecordCountButtonOnClickHandler);
    }

    /**
    * Helper method to create HTML button that is used for CreateRecord Web API Example
    * @param buttonLabel : Label for button
    * @param buttonId : ID for button
    * @param buttonValue : Value of button (attribute of button)
    * @param onClickHandler : onClick event handler to invoke for the button
    */
    private createHTMLButtonElement(buttonLabel: string, buttonId: string, buttonValue: string | null, onClickHandler: (event?: any) => void): HTMLButtonElement
    {
        const button: HTMLButtonElement = document.createElement('button');
        button.innerHTML = buttonLabel;
        button.id = buttonId;
        button.classList.add('SampleControl_WebAPI_ButtonClass');
        button.addEventListener('click', onClickHandler);

        if (buttonValue) {
            button.setAttribute('buttonvalue', buttonValue);
        }

        return button;
    }

    /**
    * Helper method to create HTML Div Element
    * @param elementClassName : Class name of div element
    * @param isHeader : True if 'header' div - adds extra class and post-fix to ID for header elements
    * @param innerText : innerText of Div Element
    */
    private createHTMLDivElement(elementClassName: string, isHeader: boolean, innerText?: string): HTMLDivElement
    {
        const div: HTMLDivElement = document.createElement('div');

        if (isHeader) {
            div.classList.add('SampleControl_WebAPI_Header');
            elementClassName += '_header';
        }

        if (innerText) {
            div.innerText = innerText.toUpperCase();
        }

        div.classList.add(elementClassName);
        return div;
    }

    /** 
    * Renders a 'result container' div element to inject the status of the example Web API calls 
    */
    private renderResultsDiv(): void
    {
        const resultDivHeader: HTMLDivElement = this.createHTMLDivElement('result_container', true, 'Result of last action');
        this._container.appendChild(resultDivHeader);

        this._resultContainerDiv = this.createHTMLDivElement('result_container', false, undefined);
        this._container.appendChild(this._resultContainerDiv);

        this.updateResultContainerText('Web API sample custom control loaded!');
    }

    /**
    * Helper method to inject HTML into result container div
    * @param statusHTML : HTML to inject into result container
    */
    private updateResultContainerText(statusHTML: string): void
    {
        if (this._resultContainerDiv) {
            this._resultContainerDiv.innerHTML = statusHTML;
        }
    }

    /**
    * Helper method to inject error string into result container div after failed Web API call
    * @param errorResponse : error object from rejected promise
    */
    private updateResultContainerTextWithErrorResponse(errorResponse: any): void
    {
        if (this._resultContainerDiv) {
            const errorHTML: string = `Error with Web API call:<br>${errorResponse.message}`;
            this._resultContainerDiv.innerHTML = errorHTML;
        }
    }

    /**
    * Helper method to generate Label for Create Buttons
    * @param entityNumber : value to set _currencyAttributeNameFriendlyName field to for this button
    */
    private getCreateRecordButtonlabel(entityNumber: string): string
    {
        return `Create record with ${TSWebAPI._currencyAttributeNameFriendlyName} of ${entityNumber}`;
    }

    /**
    * Helper method to generate ID for Create button
    * @param entityNumber : value to set _currencyAttributeNameFriendlyName field to for this button
    */
    private getCreateButtonId(entityNumber: string): string
    {
        return `create_button_${entityNumber}`;
    }

    /**
    * Event Handler for onClick of create record button
    * @param event : click event
    */
    private async createButtonOnClickHandler(event: Event): Promise<void>
    {
        const currencyAttributeValue: Number = parseInt((event.currentTarget! as Element)!.attributes.getNamedItem('buttonvalue')!.value);
        const recordName: string = `${TSWebAPI._requiredAttributeValue}_${Date.now()}`;

        let data: any = {};
        data[TSWebAPI._requiredAttributeName] = recordName;
        data[TSWebAPI._currencyAttributeName] = currencyAttributeValue;

        try {
            const response: ComponentFramework.LookupValue = await this._context.webAPI.createRecord(TSWebAPI._entityName, data);
            const id: string = response.id;
            const resultHtml: string = `Created new ${TSWebAPI._entityName} record with below values:<br><br>
                id: ${id}<br><br>
                ${TSWebAPI._requiredAttributeName}: ${recordName};<br><br>
                ${TSWebAPI._currencyAttributeName}: ${currencyAttributeValue};`
            this.updateResultContainerText(resultHtml);
        } catch (errorResponse: any) {
            this.updateResultContainerTextWithErrorResponse(errorResponse);
        }
    }

    /**
    * Event Handler for onClick of delete record button
    * @param event : click event
    */
    private async deleteButtonOnClickHandler(event: Event): Promise<void>
    {
        const lookUpOptions: ComponentFramework.UtilityApi.LookupOptions = { entityTypes: [TSWebAPI._entityName] };
        const lookUpPromise: Promise<ComponentFramework.LookupValue[]> = this._context.utils.lookupObjects(lookUpOptions);

        try {
            const data: ComponentFramework.LookupValue[] = await lookUpPromise;
            if (data && data[0]) {
                const id: string = data[0].id;
                const entityType: string = data[0].entityType;

                const response: ComponentFramework.LookupValue = await this._context.webAPI.deleteRecord(entityType, id);
                const responseId: string = response.id;
                const responseEntityType: string = response.name!;
                this.updateResultContainerText(`Deleted ${responseEntityType} record with ID: ${responseId}`);
            }
        } catch (errorResponse: any) {
            this.updateResultContainerTextWithErrorResponse(errorResponse);
        }
    }

    /**
    * Event Handler for onClick of calculate average value button
    * @param event : click event
    */
    private async calculateAverageButtonOnClickHandler(event: Event): Promise<void>
    {
        const fetchXml: string = `
            <fetch distinct="false" mapping="logical" aggregate="true">
                <entity name="${TSWebAPI._entityName}">
                    <attribute name="${TSWebAPI._currencyAttributeName}" aggregate="avg" alias="average_val"/>
                    <filter>
                        <condition attribute="${TSWebAPI._currencyAttributeName}" operator="not-null"/>
                    </filter>
                </entity>
            </fetch>
        `;

        try {
            const response: ComponentFramework.WebApi.RetrieveMultipleResponse = await this._context.webAPI.retrieveMultipleRecords(TSWebAPI._entityName, `?fetchXml=${fetchXml}`);
            const averageVal: Number = response.entities[0].average_val;
            const resultHtml: string = `Average value of ${TSWebAPI._currencyAttributeNameFriendlyName} attribute for all ${TSWebAPI._entityName} records: ${averageVal}`;
            this.updateResultContainerText(resultHtml);
        } catch (errorResponse: any) {
            this.updateResultContainerTextWithErrorResponse(errorResponse);
        }
    }

    /**
    * Event Handler for onClick of calculate record count button
    * @param event : click event
    */
    private async refreshRecordCountButtonOnClickHandler(event: Event): Promise<void>
    {
        const queryString: string = `?$select=${TSWebAPI._currencyAttributeName}&$filter=contains(${TSWebAPI._requiredAttributeName},'${TSWebAPI._requiredAttributeValue}')`;

        try {
            const response: ComponentFramework.WebApi.RetrieveMultipleResponse = await this._context.webAPI.retrieveMultipleRecords(TSWebAPI._entityName, queryString);
            
            let count1: number = 0;
            let count2: number = 0;
            let count3: number = 0;

            for (const entity of response.entities) {
                const value: number = entity[TSWebAPI._currencyAttributeName];
                if (value === 100) count1++;
                else if (value === 200) count2++;
                else if (value === 300) count3++;
            }

            const innerHTML: string = `Use above buttons to create or delete a record to see count update<br><br>
                Count of ${TSWebAPI._entityName} records with ${TSWebAPI._currencyAttributeName} of 100: ${count1};<br>
                Count of ${TSWebAPI._entityName} records with ${TSWebAPI._currencyAttributeName} of 200: ${count2};<br>
                Count of ${TSWebAPI._entityName} records with ${TSWebAPI._currencyAttributeName} of 300: ${count3};`;
            
            if (this._oDataStatusContainerDiv) {
                this._oDataStatusContainerDiv.innerHTML = innerHTML;
            }

            this.updateResultContainerText(innerHTML);
        } catch (errorResponse: any) {
            this.updateResultContainerTextWithErrorResponse(errorResponse);
        }
    }

    /**
    * Renders example use of CreateRecord Web API
    */
    private renderCreateExample(): void
    {
        const headerDiv: HTMLDivElement = this.createHTMLDivElement('create_container', true, `Click to create ${TSWebAPI._entityName} record`);
        this._container.appendChild(headerDiv);

        const value1: string = '100';
        this._createEntity1Button = this.createHTMLButtonElement(this.getCreateRecordButtonlabel(value1), this.getCreateButtonId(value1), value1, this.createButtonOnClickHandler.bind(this));

        const value2: string = '200';
        this._createEntity2Button = this.createHTMLButtonElement(this.getCreateRecordButtonlabel(value2), this.getCreateButtonId(value2), value2, this.createButtonOnClickHandler.bind(this));

        const value3: string = '300';
        this._createEntity3Button = this.createHTMLButtonElement(this.getCreateRecordButtonlabel(value3), this.getCreateButtonId(value3), value3, this.createButtonOnClickHandler.bind(this));

        this._container.appendChild(this._createEntity1Button);
        this._container.appendChild(this._createEntity2Button);
        this._container.appendChild(this._createEntity3Button);
    }

    /** 
    * Renders example use of DeleteRecord Web API
    */
    private renderDeleteExample(): void
    {
        const headerDiv: HTMLDivElement = this.createHTMLDivElement('delete_container', true, `Click to delete ${TSWebAPI._entityName} record`);
        this._deleteRecordButton = this.createHTMLButtonElement('Select record to delete', 'delete_button', null, this.deleteButtonOnClickHandler.bind(this));

        this._container.appendChild(headerDiv);
        this._container.appendChild(this._deleteRecordButton);
    }

    /** 
    * Renders example use of RetrieveMultiple Web API with Fetch XML
    */
    private renderFetchXmlRetrieveMultipleExample(): void
    {
        const containerName: string = 'fetchxml_status_container';
        const statusDivHeader: HTMLDivElement = this.createHTMLDivElement(containerName, true, `Click to calculate average value of ${TSWebAPI._currencyAttributeNameFriendlyName}`);
        const statusDiv: HTMLDivElement = this.createHTMLDivElement(containerName, false, undefined);

        this._fetchXmlRefreshButton = this.createHTMLButtonElement(`Calculate average value of ${TSWebAPI._currencyAttributeNameFriendlyName}`, 'odata_refresh',null, this.calculateAverageButtonOnClickHandler.bind(this));

        this._container.appendChild(statusDivHeader);
        this._container.appendChild(statusDiv);
        this._container.appendChild(this._fetchXmlRefreshButton);
    }

    /** 
    * Renders example use of RetrieveMultiple Web API with OData
    */
    private renderODataRetrieveMultipleExample(): void
    {
        const containerClassName: string = 'odata_status_container';

        const statusDivHeader: HTMLDivElement = this.createHTMLDivElement(containerClassName, true, 'Click to refresh record count');
        this._oDataStatusContainerDiv = this.createHTMLDivElement(containerClassName, false, undefined);

        this._oDataRefreshButton = this.createHTMLButtonElement('Refresh record count', 'odata_refresh', null, this.refreshRecordCountButtonOnClickHandler.bind(this));

        this._container.appendChild(statusDivHeader);
        this._container.appendChild(this._oDataStatusContainerDiv);
        this._container.appendChild(this._oDataRefreshButton);
    }
}
